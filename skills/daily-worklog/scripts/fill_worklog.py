import argparse
import json
import os
import random
import subprocess
import sys
import time
import urllib.request
from datetime import datetime, timedelta

import yaml
from playwright.sync_api import sync_playwright
from engine import (
    ASSETS_DIR,
    load_task_db, match_tasks, aggregate_tasks, filter_by_date_rules,
    replace_month_in_logs, is_project_specific_keyword, adjust_to_eight, round_to_half,
    weighted_random_tasks, supplement_with_random, parse_manual_tasks, append_to_task_db,
    is_quality_planning_related, should_allow_quality_planning, merge_project_rows,
    get_projects_with_quality_planning, enforce_daily_quality_planning_limit,
    infer_workitem_levels, _generate_month_end_tasks, _get_project_core_keyword,
    _extract_feature_signatures, _has_signature_overlap, _safe_str,
)

# 读取配置
try:
    config_path = ASSETS_DIR / "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
except Exception:
    config = {}
default_department = config.get("preferences", {}).get("default_department", "电信与AIoT业务线")


BASE_URL = "https://prj.chinasoftinc.com/prj/index.jsp#"
CDP_URL = "http://127.0.0.1:9222"

# 工作内容关键词 → 工作项智能映射表
_KEYWORD_TO_WORKITEM = [
    ("审计", "审计"),
    ("流程遵从度", "审计"),
    ("投标", "招投标管理"),
    ("招投标", "招投标管理"),
    ("入围供应商", "招投标管理"),
    ("质量策划", "制定与辅导项目质量策划"),
    ("策划评审", "制定与辅导项目质量策划"),
    ("自动化测试", "AI工具落地与应用支持"),
    ("AI测试", "AI工具落地与应用支持"),
    ("效果查看", "AI工具落地与应用支持"),
    ("AI工具", "AI工具落地与应用支持"),
    ("汇报", "AI工具引入与推广策划"),
    ("思路对齐", "AI工具引入与推广策划"),
    ("推广", "AI工具引入与推广策划"),
    ("培训", "AI工具引入与推广策划"),
    ("CP评审", "关键节点评审与过程审视"),
    ("评审", "关键节点评审与过程审视"),
    ("SOW", "SOW评审及问题闭环"),
    ("验收", "预验收与验收管理"),
    ("任命", "项目立项与排产计划管理"),
    ("备案", "内部专项工作"),
    ("度量", "度量分析"),
    ("成本", "项目成本监控"),
    ("监控", "项目成本监控"),
    ("风险", "项目风险识别"),
]


def _infer_workitem_from_log(log_text: str) -> str | None:
    """根据工作内容描述推断对应的工作项叶子节点。"""
    if not log_text:
        return None
    text = log_text.lower()
    for kw, workitem in _KEYWORD_TO_WORKITEM:
        if kw in text:
            return workitem
    return None


def screenshot(page, name: str):
    path = f"screenshot_{name}.png"
    try:
        # 跳过字体等待，直接截图视口
        page.evaluate("document.fonts.ready")
        page.screenshot(path=path, full_page=False, timeout=5000)
        print(f"[截图] {path}")
    except Exception as e:
        print(f"[截图失败] {path}: {e}")
    return path


def ensure_on_calendar(page):
    """确保当前页面在日历视图。无论当前在什么页面，强制回到日历首页，防止停留在详情页填错日期。"""
    print(f"当前页面 URL: {page.url}")
    if "prj.chinasoftinc.com/prj/index.jsp" in page.url:
        print("已在目标页面，跳过导航")
    else:
        page.goto(BASE_URL)
    page.wait_for_timeout(3000)
    print("已导航到日历首页")


def get_target_frame(page):
    """找到包含日历/填报表格的 iframe（URL 含 workreport/toReport）"""
    for frame in page.frames:
        if "workreport/toReport" in frame.url:
            print(f"找到目标 iframe: {frame.url}")
            return frame
    return None


def is_date_filled(frame, target_day: int) -> bool:
    """检查目标日期是否已有填报记录（单元格中包含'申报'等字样）"""
    js_code = r"""
    () => {
        const target = '__DAY__';
        const selectors = ['td', 'div', 'span', 'li', 'a'];
        let debugInfo = [];
        for (const sel of selectors) {
            for (const el of document.querySelectorAll(sel)) {
                const text = el.innerText.trim();
                const lines = text.split('\n').map(l => l.trim()).filter(l => l);
                if (lines[0] === target) {
                    const hasRecord = /申报|审批|草稿|已填|不通过/.test(text);
                    debugInfo.push({selector: sel, text: text.substring(0, 80), hasRecord});
                    if (hasRecord) return {filled: true, debug: debugInfo};
                }
            }
        }
        return {filled: false, debug: debugInfo};
    }
    """.replace('__DAY__', str(target_day))
    result = frame.evaluate(js_code)
    # 兼容旧返回值（直接 bool）
    if isinstance(result, bool):
        return result
    debug = result.get('debug', [])
    if debug:
        for d in debug[:3]:
            print(f"[调试] 日期检测: {d['text']} -> hasRecord={d['hasRecord']}")
    return result.get('filled', False)


def is_detail_page_has_draft(frame) -> bool:
    """在详情页检测是否已有填报记录（不依赖日历视图的文本标记）。
    注意：避免被'保存草稿'按钮文字误报，优先检查实际数据。"""
    js_code = """
    () => {
        // 1. 检查 datagrid 行中是否有实际填写数据
        const rows = document.querySelectorAll('.datagrid-btable tr');
        for (const row of rows) {
            const inputs = row.querySelectorAll('input');
            for (const inp of inputs) {
                const val = (inp.value || '').trim();
                const ph = (inp.getAttribute('placeholder') || '').trim();
                if (val && val !== '' && val !== '0' && !ph) {
                    return {filled: true, reason: '表格中有已填写数据'};
                }
            }
        }
        // 2. 检查 textarea 是否有内容
        const textareas = document.querySelectorAll('textarea');
        for (const ta of textareas) {
            const val = (ta.value || '').trim();
            if (val && val.length > 5) {
                return {filled: true, reason: '日志框有内容'};
            }
        }
        // 3. 检查 combo 隐藏值（项目/工作项是否已选）
        const combos = document.querySelectorAll('.textbox-value');
        for (const combo of combos) {
            const val = (combo.value || '').trim();
            if (val && val !== '' && !isNaN(Number(val))) {
                return {filled: true, reason: '下拉框有选中值'};
            }
        }
        return {filled: false, reason: '无数据'};
    }
    """
    result = frame.evaluate(js_code)
    if isinstance(result, dict):
        print(f"[调试] 详情页检测: {result.get('reason')}")
        return result.get('filled', False)
    return False


def find_and_click_date(frame, target_day: int) -> bool:
    """在日历中找到并点击目标日期，点击后验证是否真正进入详情页（出现保存草稿按钮或工时输入框）。"""
    js_code = r"""
    () => {
        // 先统计页面元素
        const stats = {
            tables: document.querySelectorAll('table').length,
            tds: document.querySelectorAll('td').length,
            divs: document.querySelectorAll('div').length,
            spans: document.querySelectorAll('span').length
        };

        // 策略1：查找所有 td（不限制父级）
        const allTds = document.querySelectorAll('td');
        for (const cell of allTds) {
            const text = cell.innerText.trim();
            const lines = text.split('\n').map(l => l.trim());
            if (lines[0] === '__DAY__') {
                cell.click();
                return { success: true, clicked: text, strategy: 'td', stats: stats };
            }
        }

        // 策略2：查找所有 div 和 span，匹配纯数字文本
        const allElements = document.querySelectorAll('div, span, li, a');
        for (const el of allElements) {
            const text = el.innerText.trim();
            const lines = text.split('\n').map(l => l.trim());
            if (lines[0] === '__DAY__' && el.children.length <= 3) {
                el.click();
                return { success: true, clicked: text, strategy: 'div/span', stats: stats };
            }
        }

        // 策略3：查找包含目标数字的元素
        const allEls = document.querySelectorAll('*');
        let debugInfo = [];
        for (const el of allEls) {
            const text = el.innerText.trim();
            const lines = text.split('\n').map(l => l.trim());
            if (lines[0] && lines[0].match(/^\d+$/)) {
                debugInfo.push(el.tagName + ':' + lines[0]);
            }
            if (lines[0] === '__DAY__') {
                el.click();
                return { success: true, clicked: text, strategy: 'fallback', stats: stats };
            }
        }

        // 策略4：打印所有包含目标数字的元素（innerText）
        let allMatches = [];
        for (const el of document.querySelectorAll('*')) {
            const text = (el.innerText || '').trim();
            if (text.includes('__DAY__')) {
                allMatches.push('IT:' + el.tagName + ':' + text.replace(/\n/g, '|').substring(0, 60));
            }
        }

        // 策略5：用 textContent 搜索（包含 Shadow DOM 和隐藏文本）
        let tcMatches = [];
        for (const el of document.querySelectorAll('*')) {
            const text = (el.textContent || '').trim();
            if (text.includes('__DAY__')) {
                tcMatches.push('TC:' + el.tagName + ':' + text.replace(/\n/g, '|').substring(0, 60));
            }
        }

        // 策略6：搜索 data- 属性和 title 属性
        let attrMatches = [];
        for (const el of document.querySelectorAll('*')) {
            for (const attr of el.attributes) {
                if (attr.value.includes('__DAY__')) {
                    attrMatches.push('ATTR:' + el.tagName + ':' + attr.name + '=' + attr.value.substring(0, 40));
                }
            }
        }

        // 策略7：HTML 中是否包含目标数字
        const htmlHas = document.documentElement.outerHTML.includes('__DAY__');

        return { success: false, stats: stats, debug: debugInfo.slice(0, 50), allMatches: allMatches.slice(0, 20), tcMatches: tcMatches.slice(0, 20), attrMatches: attrMatches.slice(0, 20), htmlHas: htmlHas, reason: 'no matching cell' };
    }
    """.replace('__DAY__', str(target_day))
    result = frame.evaluate(js_code)
    if not result or not result.get('success'):
        print(f"[调试] 查找日期 {target_day} 失败: {result}")
        return False

    # 点击后等待页面跳转/加载，然后验证是否进入详情页
    frame.wait_for_timeout(2000)
    js_verify = """() => {
        const hasDraftBtn = !!document.getElementById('btnSaveDraft');
        const hasHourInput = document.querySelectorAll("input[type='number']").length > 0;
        const hasCalendar = document.querySelectorAll('td').length > 50;
        return { inDetailPage: hasDraftBtn || hasHourInput, hasCalendar: hasCalendar, draftBtn: hasDraftBtn, hourInput: hasHourInput };
    }"""
    verify = frame.evaluate(js_verify)
    print(f"[调试] 点击日期后页面状态: {verify}")
    if verify and not verify.get('inDetailPage'):
        print(f"[错误] 点击日期 {target_day} 后未进入详情页，系统可能禁止选择该日期（未来日期或休息日）")
        return False
    return True


def add_row_in_section(frame, section_title: str):
    """
    在指定区域（项目工时填报 / 公共工时填报）点击 '+ 增加行'。
    策略：先找到区域标题，再取该区域下方的第一个 '+ 增加行' 按钮。
    """
    if "项目" in section_title:
        btn = frame.locator("button:has-text('+ 增加行'), .el-button:has-text('+ 增加行')").first
    else:
        btns = frame.locator("button:has-text('+ 增加行'), .el-button:has-text('+ 增加行')").all()
        btn = btns[1] if len(btns) >= 2 else (btns[-1] if btns else None)
    if btn:
        btn.evaluate("el => el.click()")
        frame.wait_for_timeout(1500)


def select_project(frame, project_name: str):
    """
    在最新添加的项目工时行中搜索并选择项目。
    使用 JS 操作 EasyUI combo 组件：通过 value/placeholder 定位 input，
    调用 combo('showPanel') 或点击 combo-arrow 展开面板，再填充搜索框。
    """

    # 1. 展开项目选择下拉面板
    js_open = """
    () => {
        // 找到项目工时填报区域
        var sections = document.querySelectorAll('div, section');
        var projectSection = null;
        for (var i = 0; i < sections.length; i++) {
            var text = sections[i].innerText || '';
            if (text.includes('项目工时填报')) {
                projectSection = sections[i];
                break;
            }
        }
        if (!projectSection) return { success: false, reason: '项目工时区域未找到' };

        // 从后往前找最后一个 value/placeholder 为"请选择项目"的 input
        var inputs = projectSection.querySelectorAll('input');
        var targetInput = null;
        for (var i = inputs.length - 1; i >= 0; i--) {
            var inp = inputs[i];
            var val = (inp.value || '').trim();
            var ph  = (inp.getAttribute('placeholder') || '').trim();
            if (val === '请选择项目' || ph === '请选择项目') {
                targetInput = inp;
                break;
            }
        }
        if (!targetInput) return { success: false, reason: '未找到请选择项目输入框' };

        // 优先使用 EasyUI API 展开面板
        if (window.jQuery) {
            var $combo = jQuery(targetInput).closest('.textbox.combo, .combo');
            if ($combo.length) {
                try {
                    $combo.combo('showPanel');
                    return { success: true, method: 'comboAPI' };
                } catch(e) {}
            }
        }

        // 兜底：点击下拉箭头
        var combo = targetInput.closest('.textbox.combo, .combo');
        if (combo) {
            var arrow = combo.querySelector('.combo-arrow');
            if (arrow) {
                arrow.click();
                return { success: true, method: 'arrowClick' };
            }
        }

        // 最后兜底：直接点击 input
        targetInput.click();
        return { success: true, method: 'inputClick' };
    }
    """
    result = frame.evaluate(js_open)
    print(f"[调试] 展开项目选择: {result}")
    if not result or not result.get('success'):
        print(f"[错误] 展开项目选择失败: {result}")
        return False
    frame.wait_for_timeout(1500)

    # 2. 在弹出的面板中查找输入框并填充项目名称
    js_fill = f"""
    () => {{
        var name = {json.dumps(project_name)};
        var isReadonlyOrDisplay = function(inp) {{
            return inp.classList.contains('validatebox-readonly') || inp.classList.contains('textbox-prompt');
        }};

        // 2.1 在可见面板中查找搜索框（先按placeholder匹配，再回退到任意text输入框）
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        var targetInput = null;
        var debugInfo = {{ panelCount: panels.length, panelInputs: [] }};
        for (var i = 0; i < panels.length; i++) {{
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {{
                // 扫描该面板内所有input用于诊断
                var allInps = p.querySelectorAll('input');
                var scan = [];
                for (var k = 0; k < allInps.length; k++) {{
                    var inp = allInps[k];
                    var ir = inp.getBoundingClientRect();
                    scan.push({{
                        type: inp.type, className: inp.className, placeholder: inp.placeholder,
                        readonly: inp.readOnly, width: ir.width, height: ir.height
                    }});
                }}
                debugInfo.panelInputs.push(scan);

                // 策略A：按placeholder匹配（面板内不排除 wr-prj-toolbar-input，它就是搜索框）
                var inputs = p.querySelectorAll("input[placeholder*='项目名称']");
                for (var j = 0; j < inputs.length; j++) {{
                    var inp = inputs[j];
                    var ir = inp.getBoundingClientRect();
                    if (ir.width > 0 && ir.height > 0 && !isReadonlyOrDisplay(inp)) {{
                        targetInput = inp;
                        break;
                    }}
                }}
                if (targetInput) break;
                // 策略B：回退到面板内任意可见的text输入框（排除datagrid内部可编辑单元格）
                var textInputs = p.querySelectorAll("input[type='text']");
                for (var j = 0; j < textInputs.length; j++) {{
                    var inp = textInputs[j];
                    var ir = inp.getBoundingClientRect();
                    var isGridCell = inp.classList.contains('datagrid-editable-input');
                    if (ir.width > 0 && ir.height > 0 && !isReadonlyOrDisplay(inp) && !isGridCell) {{
                        targetInput = inp;
                        break;
                    }}
                }}
                if (targetInput) break;
            }}
        }}

        // 2.2 全局回退：先按placeholder，再按任意text输入框（同样排除readonly）
        if (!targetInput) {{
            var allInputs = document.querySelectorAll("input[placeholder*='项目名称']");
            for (var i = 0; i < allInputs.length; i++) {{
                var inp = allInputs[i];
                var ir = inp.getBoundingClientRect();
                if (ir.width > 0 && ir.height > 0 && !inp.classList.contains('wr-prj-toolbar-input') && !isReadonlyOrDisplay(inp)) {{
                    targetInput = inp;
                    break;
                }}
            }}
        }}
        if (!targetInput) {{
            var allTextInputs = document.querySelectorAll("input[type='text']");
            for (var i = 0; i < allTextInputs.length; i++) {{
                var inp = allTextInputs[i];
                var ir = inp.getBoundingClientRect();
                if (ir.width > 0 && ir.height > 0 && !inp.classList.contains('wr-prj-toolbar-input') && !isReadonlyOrDisplay(inp)) {{
                    targetInput = inp;
                    break;
                }}
            }}
        }}

        if (!targetInput) return {{ success: false, reason: '未找到项目名称输入框', debug: debugInfo }};

        // 通过多种方式确保 EasyUI/datagrid 能感知到值变化
        if (window.jQuery && jQuery.fn && jQuery.fn.val) {{
            jQuery(targetInput).val(name).trigger('input').trigger('change');
        }}
        targetInput.value = name;
        targetInput.focus();
        targetInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
        targetInput.dispatchEvent(new KeyboardEvent('keydown', {{ bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13 }}));
        targetInput.dispatchEvent(new KeyboardEvent('keypress', {{ bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13 }}));
        targetInput.dispatchEvent(new KeyboardEvent('keyup', {{ bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13 }}));
        targetInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        targetInput.blur();
        return {{ success: true, className: targetInput.className, placeholder: targetInput.placeholder }};
    }}
    """
    result = frame.evaluate(js_fill)
    print(f"[调试] 填充项目名称: {result}")
    if not result or not result.get('success'):
        print(f"[错误] 填充项目名称失败: {result}")
        return False
    frame.wait_for_timeout(800)

    # 3. 点击查询按钮
    js_query = """
    () => {
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        var targetBtn = null;
        for (var i = 0; i < panels.length; i++) {
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {
                var btns = p.querySelectorAll('button, a, span, div, input[type="button"]');
                for (var j = 0; j < btns.length; j++) {
                    var b = btns[j];
                    if ((b.innerText || b.textContent || '').trim() === '查询') {
                        var br = b.getBoundingClientRect();
                        if (br.width > 0 && br.height > 0) {
                            targetBtn = b;
                            break;
                        }
                    }
                }
            }
            if (targetBtn) break;
        }
        if (targetBtn) {
            targetBtn.click();
            return { success: true };
        }
        return { success: false, reason: '未找到查询按钮' };
    }
    """
    result = frame.evaluate(js_query)
    print(f"[调试] 点击查询: {result}")
    frame.wait_for_timeout(1500)

    # 4. 点击结果表格第一行（限制在项目选择面板内，并输出点击行的文本用于诊断）
    js_select_row = rf"""
    () => {{
        var name = {json.dumps(project_name)};
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        for (var i = 0; i < panels.length; i++) {{
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {{
                var rows = p.querySelectorAll('.datagrid-row');
                var shortName = name.replace(/\d{{4}}[-~]\d{{4}}年?/g, '').trim();
                for (var j = 0; j < rows.length; j++) {{
                    var r = rows[j];
                    var rr = r.getBoundingClientRect();
                    var rowText = (r.innerText || r.textContent || '').trim();
                    // 优先点击包含目标项目名称（或去掉年份前缀后的短名）的行
                    if (rr.width > 0 && rr.height > 0 && (rowText.indexOf(name) >= 0 || rowText.indexOf(shortName) >= 0)) {{
                        r.click();
                        return {{ success: true, rowText: rowText.substring(0, 80), rowIndex: j }};
                    }}
                }}
                // 无匹配行时不回退，避免误选错误项目
            }}
        }}
        return {{ success: false, reason: '未找到结果行' }};
    }}
    """
    result = frame.evaluate(js_select_row)
    print(f"[调试] 选择结果行: {result}")
    frame.wait_for_timeout(800)

    # 5. 验证回填：只检查 combo 显示框/隐藏值框，排除搜索框(wr-prj-toolbar-input)
    js_verify = f"""
    () => {{
        var name = {json.dumps(project_name)};
        var combos = document.querySelectorAll('.textbox.combo');
        for (var i = 0; i < combos.length; i++) {{
            var combo = combos[i];
            var textInput = combo.querySelector('input.textbox-text');
            var valueInput = combo.querySelector('input.textbox-value');
            var val = (textInput && textInput.value || '').trim();
            var hid = (valueInput && valueInput.value || '').trim();
            if (val && val !== '请选择项目' && (val === name || val.indexOf(name) >= 0 || name.indexOf(val) >= 0)) {{
                return {{ filled: true, value: val, hidden: hid, className: textInput.className }};
            }}
        }}
        return {{ filled: false }};
    }}
    """
    verify = frame.evaluate(js_verify)
    print(f"[调试] 项目回填验证: {verify}")
    if verify and verify.get('filled'):
        filled_name = verify.get('value') or project_name
        print(f"[成功] 项目已正确回填: {filled_name}")
        close_all_easyui_panels(frame)
        return filled_name

    # 6. 兜底：直接写入 combo 显示框和隐藏值框
    print("[警告] 下拉流程未成功回填，尝试直接写入 combo 值...")
    _direct_fill_combo_text(frame, "请选择项目", project_name, "项目")
    close_all_easyui_panels(frame)
    return project_name


def close_all_easyui_panels(frame):
    """温和关闭所有已打开的 EasyUI 下拉面板及遮罩层，避免破坏内部状态。"""
    frame.evaluate("""
    () => {
        if (window.jQuery) {
            jQuery('.textbox.combo').each(function() {
                try { jQuery(this).combo('hidePanel'); } catch(e) {}
            });
        }
        jQuery('.combo-panel:visible, .window-mask').each(function() {
            jQuery(this).hide();
        });
    }
    """)
    frame.wait_for_timeout(800)


def _set_combo_value_by_api(frame, target_text: str):
    """
    通过 EasyUI jQuery API 直接设置下拉值。
    当模拟点击无效时，此方法是最终兜底方案。
    支持 combogrid、combotreegrid、combobox。
    """
    js = f"""
    () => {{
        var inputs = document.querySelectorAll("input[placeholder*='工作项']");
        if (inputs.length === 0) return {{ error: 'no input found' }};
        var lastInput = inputs[inputs.length - 1];
        if (!window.jQuery) return {{ error: 'jQuery not available' }};
        var $input = jQuery(lastInput);

        // 1. 尝试 combogrid（下拉表格，面板内有 datagrid）
        if ($input.data('combogrid')) {{
            try {{
                var opts = $input.combogrid('options');
                var grid = $input.combogrid('grid');
                var rows = grid.datagrid('getRows');
                var targetRow = null;
                for (var i = 0; i < rows.length; i++) {{
                    var r = rows[i];
                    var txt = r[opts.textField || 'text'] || '';
                    if (txt === '{target_text}' || txt.indexOf('{target_text}') >= 0) {{
                        targetRow = r;
                        break;
                    }}
                }}
                if (!targetRow) return {{ error: 'row not found in combogrid', rowsCount: rows.length }};
                $input.combogrid('setValue', targetRow[opts.idField || 'id']);
                return {{ success: true, type: 'combogrid', value: targetRow[opts.idField], text: targetRow[opts.textField] }};
            }} catch(e) {{
                return {{ error: 'combogrid failed: ' + e.message }};
            }}
        }}

        // 2. 尝试 combotreegrid（树形表格）
        if ($input.data('combotreegrid')) {{
            try {{
                var opts = $input.combotreegrid('options');
                var grid = $input.combotreegrid('grid');
                var allRows = grid.datagrid('getRows');
                function search(rows) {{
                    for (var i = 0; i < rows.length; i++) {{
                        var r = rows[i];
                        var txt = r[opts.textField || opts.treeField || 'text'] || '';
                        if (txt === '{target_text}' || txt.indexOf('{target_text}') >= 0) return r;
                        if (r.children && r.children.length) {{
                            var found = search(r.children);
                            if (found) return found;
                        }}
                    }}
                    return null;
                }}
                var targetRow = search(allRows);
                if (!targetRow) return {{ error: 'row not found in combotreegrid' }};
                $input.combotreegrid('setValue', targetRow[opts.idField || 'id']);
                return {{ success: true, type: 'combotreegrid', value: targetRow[opts.idField], text: targetRow[opts.textField] }};
            }} catch(e) {{
                return {{ error: 'combotreegrid failed: ' + e.message }};
            }}
        }}

        // 3. 尝试 combobox
        if ($input.data('combobox')) {{
            try {{
                var data = $input.combobox('getData');
                var target = null;
                for (var i = 0; i < data.length; i++) {{
                    var item = data[i];
                    var txt = item.text || '';
                    if (txt === '{target_text}' || txt.indexOf('{target_text}') >= 0) {{
                        target = item;
                        break;
                    }}
                }}
                if (!target) return {{ error: 'item not found in combobox', dataCount: data.length }};
                $input.combobox('setValue', target.value || target.id);
                return {{ success: true, type: 'combobox', value: target.value, text: target.text }};
            }} catch(e) {{
                return {{ error: 'combobox failed: ' + e.message }};
            }}
        }}

        return {{ error: 'unknown combo type', dataKeys: Object.keys($input.data()) }};
    }}
    """
    return frame.evaluate(js)


def _expand_all_tree_nodes(frame):
    """展开下拉面板中所有折叠的树节点，确保叶子节点可见。"""
    frame.evaluate("""
    () => {
        var hits = document.querySelectorAll('.tree-collapsed, .tree-node-collapsed');
        for (var i = 0; i < hits.length; i++) {
            hits[i].click();
        }
    }
    """)
    frame.wait_for_timeout(800)


def _select_from_easyui_panel(frame, target_text: str, label: str, leaf_only: bool = True):
    """
    在可见的 EasyUI 下拉面板中点击匹配文本。
    增强版：先展开所有折叠节点，再用 JS 在面板内通用查找包含目标文本的元素，
    优先点击叶子节点，回退到 Playwright 真实点击或 JS 点击。
    leaf_only=False 时允许选择非叶子节点（用于部门等业务线父节点选择）。
    """
    # 展开所有折叠的树节点，确保内容可见
    _expand_all_tree_nodes(frame)
    frame.wait_for_timeout(500)

    # 调试：打印所有可见面板的内容摘要
    js_debug = """() => {
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        var out = [];
        for (var i = 0; i < panels.length; i++) {
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {
                var treeNodes = p.querySelectorAll('.tree-node');
                var comboItems = p.querySelectorAll('.combobox-item');
                var datagridRows = p.querySelectorAll('.datagrid-row');
                var items = [];
                for (var n of treeNodes) {
                    var t = (n.innerText || n.textContent || '').trim().substring(0, 60);
                    items.push({type:'tree-node', text:t, isLeaf: n.querySelector('.tree-hit')===null});
                }
                for (var n of comboItems) {
                    var t = (n.innerText || n.textContent || '').trim().substring(0, 60);
                    items.push({type:'combobox-item', text:t});
                }
                for (var n of datagridRows) {
                    var t = (n.innerText || n.textContent || '').trim().substring(0, 60);
                    items.push({type:'datagrid-row', text:t});
                }
                out.push({itemCount: items.length, items: items.slice(0, 15)});
            }
        }
        return out;
    }"""
    debug = frame.evaluate(js_debug)
    print(f"[调试] {label} 下拉内容摘要: {json.dumps(debug, ensure_ascii=False)[:600]}")

    # 通用 JS 查找：在可见面板内精确搜索具体选项元素（tree-node / combobox-item / datagrid-row）
    js_find = f"""() => {{
        var target = '{target_text}';
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        for (var i = 0; i < panels.length; i++) {{
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {{
                var selectors = ['.tree-node', '.combobox-item', '.datagrid-row'];
                for (var s = 0; s < selectors.length; s++) {{
                    var items = p.querySelectorAll(selectors[s]);
                    for (var j = 0; j < items.length; j++) {{
                        var item = items[j];
                        var text = (item.textContent || '').trim();
                        if (text === target || text.includes(target)) {{
                            // 对 tree-node 默认只选叶子节点；leaf_only=False 时允许选择父节点
                            var isLeaf = true;
                            if (selectors[s] === '.tree-node') {{
                                isLeaf = item.querySelector('.tree-hit') === null;
                                if (!isLeaf && {str(leaf_only).lower()}) continue;
                            }}
                            var r = item.getBoundingClientRect();
                            return {{
                                found: true,
                                text: text.substring(0, 80),
                                tag: item.tagName,
                                className: item.className,
                                isLeaf: isLeaf,
                                left: r.left, top: r.top, width: r.width, height: r.height
                            }};
                        }}
                    }}
                }}
            }}
        }}
        return {{ found: false }};
    }}"""

    result = frame.evaluate(js_find)
    print(f"[调试] {label} 查找 '{target_text}' 结果: {json.dumps(result, ensure_ascii=False)[:300]}")

    if not result or not result.get('found'):
        # 收集所有可见选项（叶子节点优先）供用户参考
        js_list = """() => {
            var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
            var leafTexts = [];
            var allTexts = [];
            for (var i = 0; i < panels.length; i++) {
                var p = panels[i];
                var rect = p.getBoundingClientRect();
                if (rect.width > 0 && rect.height > 0) {
                    var treeNodes = p.querySelectorAll('.tree-node');
                    for (var n of treeNodes) {
                        var t = (n.innerText || n.textContent || '').trim();
                        if (t) {
                            var isLeaf = n.querySelector('.tree-hit') === null;
                            if (isLeaf) leafTexts.push(t);
                            allTexts.push(t);
                        }
                    }
                    var comboItems = p.querySelectorAll('.combobox-item');
                    for (var n of comboItems) {
                        var t = (n.innerText || n.textContent || '').trim();
                        if (t) { leafTexts.push(t); allTexts.push(t); }
                    }
                }
            }
            return { leafTexts: leafTexts.slice(0, 30), allTexts: allTexts.slice(0, 30) };
        }"""
        options = frame.evaluate(js_list)
        print(f"[警告] 选择{label} '{target_text}' 失败: 未找到匹配元素")
        if options:
            print(f"[参考] 该下拉框中可用的叶子节点: {options.get('leafTexts', [])}")
            print(f"[参考] 全部可见选项: {options.get('allTexts', [])}")
        return False

    # Playwright click 已多次定位到错误元素（祖先容器），直接跳过，改用更可靠的 JS click
    print(f"[调试] 跳过 Playwright click，直接使用 JS click 选择 {label}")

    # Fallback: 纯 JS click（精确搜索具体选项元素，避免匹配到祖先容器）
    js_click = f"""() => {{
        var target = '{target_text}';
        var panels = document.querySelectorAll('.panel-htop, .combo-panel, .panel');
        for (var i = 0; i < panels.length; i++) {{
            var p = panels[i];
            var rect = p.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) {{
                var selectors = ['.tree-node', '.combobox-item', '.datagrid-row'];
                for (var s = 0; s < selectors.length; s++) {{
                    var items = p.querySelectorAll(selectors[s]);
                    for (var j = 0; j < items.length; j++) {{
                        var item = items[j];
                        var text = (item.textContent || '').trim();
                        if (text === target || text.includes(target)) {{
                            if (selectors[s] === '.tree-node') {{
                                var isLeaf = item.querySelector('.tree-hit') === null;
                                if (!isLeaf && {str(leaf_only).lower()}) continue;
                            }}
                            item.scrollIntoView({{ block: 'nearest', inline: 'nearest' }});
                            var down = new MouseEvent('mousedown', {{ bubbles: true, cancelable: true, view: window }});
                            var up = new MouseEvent('mouseup', {{ bubbles: true, cancelable: true, view: window }});
                            var clk = new MouseEvent('click', {{ bubbles: true, cancelable: true, view: window }});
                            item.dispatchEvent(down);
                            item.dispatchEvent(up);
                            item.dispatchEvent(clk);
                            return {{ success: true, clicked: selectors[s], text: text.substring(0, 80) }};
                        }}
                    }}
                }}
            }}
        }}
        return {{ success: false }};
    }}"""

    click_result = frame.evaluate(js_click)
    if click_result and click_result.get('success'):
        print(f"[调试] 选择{label} '{target_text}' 成功: JS click")
        frame.wait_for_timeout(600)
        return True

    print(f"[警告] 选择{label} '{target_text}' 失败: 所有点击策略均无效")
    return False


def _click_combo_arrow_by_placeholder(frame, keyword: str, section_hint: str = None, index: int = -1):
    """
    通过 placeholder 定位输入框，在其祖先 .textbox.combo 中点击 .combo-arrow 展开下拉。
    如果指定了 section_hint（如"项目工时填报"或"公共工时填报"），优先在对应区域内查找。
    使用面积最大的匹配区域，避免误取标题小 div。
    index 指定区域内第几个输入框（默认最后一个）。
    """
    section_filter = ""
    if section_hint:
        section_filter = f"""
        var sections = document.querySelectorAll('div, section');
        var targetSection = null;
        var maxArea = 0;
        for (var i = 0; i < sections.length; i++) {{
            var s = sections[i];
            if ((s.innerText || '').includes('{section_hint}')) {{
                var area = (s.offsetWidth || 0) * (s.offsetHeight || 0);
                if (area > maxArea) {{
                    maxArea = area;
                    targetSection = s;
                }}
            }}
        }}
        var inputs = [];
        if (targetSection) {{
            inputs = targetSection.querySelectorAll("input[placeholder*='{keyword}']");
        }}
        if (inputs.length === 0) {{
            inputs = document.querySelectorAll("input[placeholder*='{keyword}']");
        }}
        """
    else:
        section_filter = f"""var inputs = document.querySelectorAll("input[placeholder*='{keyword}']");"""

    js = f"""
    () => {{
        {section_filter}
        if (inputs.length === 0) return {{ error: 'no input found' }};
        const idx = {index} < 0 ? inputs.length + ({index}) : {index};
        const target = inputs[idx];
        if (!target) return {{ error: 'index out of range', count: inputs.length, idx: idx }};
        const combo = target.closest('.textbox.combo') || target.closest('.combo');
        if (!combo) return {{ error: 'no combo found' }};
        const arrow = combo.querySelector('.combo-arrow');
        if (arrow) {{
            arrow.click();
            return {{ success: true, idx: idx }};
        }}
        return {{ error: 'no arrow found' }};
    }}
    """
    return frame.evaluate(js)


def get_combo_value_by_placeholder(frame, keyword: str, index: int = -1):
    """读取指定 placeholder 对应输入框的当前值，用于验证下拉是否真正选中。"""
    js = f"""
    () => {{
        const inputs = document.querySelectorAll("input[placeholder*='{keyword}']");
        const idx = {index} < 0 ? inputs.length + ({index}) : {index};
        const target = inputs[idx];
        return target ? (target.value || '') : null;
    }}
    """
    return frame.evaluate(js)


def select_department(frame, dept_name: str):
    """
    选择部门（EasyUI combotree）。
    先关闭其他面板 → 点击 arrow 展开 → 等待数据加载 → 搜索并点击匹配项 → 关闭面板。
    """
    close_all_easyui_panels(frame)
    frame.wait_for_timeout(300)
    _click_combo_arrow_by_placeholder(frame, "部门")
    frame.wait_for_timeout(1500)
    _select_from_easyui_panel(frame, dept_name, "部门", leaf_only=False)
    frame.wait_for_timeout(400)
    close_all_easyui_panels(frame)


def select_workitem_dropdown(frame, workitem_name: str, row_index: int = 0) -> bool:
    """
    在项目工时区域选择工作项（EasyUI 下拉）。
    先关闭其他面板 → 点击 arrow 展开 → 等待数据加载 → 搜索并点击匹配项。
    选择后验证回填，失败则重试一次，不再兜底写入文本（避免 hidden value 格式错误导致脏数据不可删除）。
    row_index 用于精确定位第几行的工作项输入框。
    返回 True/False 表示是否成功。
    """
    for attempt in range(2):
        prefix = "首次" if attempt == 0 else "二次"
        close_all_easyui_panels(frame)
        frame.wait_for_timeout(300)
        arrow_result = _click_combo_arrow_by_placeholder(frame, "工作项", section_hint="项目工时填报", index=row_index)
        print(f"[调试] {prefix}展开工作项下拉 (第{row_index}行): {arrow_result}")
        frame.wait_for_timeout(1500)
        ok = _select_from_easyui_panel(frame, workitem_name, "工作项")
        frame.wait_for_timeout(400)

        # 回填验证
        val = _get_workitem_value(frame, "项目工时填报", row_index)
        if val == workitem_name:
            print(f"[验证] 工作项已正确回填: '{val}'")
            close_all_easyui_panels(frame)
            return True

        if not ok:
            print(f"[警告] {prefix}面板选择工作项 '{workitem_name}' 失败，当前值='{val}'")
        else:
            print(f"[警告] {prefix}面板选择成功但回填验证失败，当前值='{val}'")

    # 连续两次失败，不再兜底写入文本，避免产生不可删除的脏数据
    print(f"[错误] 工作项 '{workitem_name}' 连续两次选择失败，跳过该行。请检查任务库中该工作项名称是否正确。")
    close_all_easyui_panels(frame)
    return False


def _direct_fill_combo_text(frame, keyword: str, text: str, label: str):
    """
    直接往 EasyUI combo/combogrid 的显示输入框写入文本，并同步 hidden value input，
    然后触发 blur/change 事件让组件接受值。绕过所有下拉点击。
    """
    js = f"""
    () => {{
        var inputs = document.querySelectorAll("input[placeholder*='{keyword}']");
        if (inputs.length === 0) return {{ error: 'no input found' }};
        var lastInput = inputs[inputs.length - 1];
        var combo = lastInput.closest('.textbox.combo');

        // 写入显示文本
        lastInput.value = '{text}';

        // 同步 hidden value input（如果存在）
        var valueInput = combo ? combo.querySelector('input.textbox-value') : null;
        if (valueInput) {{
            valueInput.value = '{text}';
        }}

        // 触发 change，让 EasyUI 识别（避免 blur 重新打开面板）
        if (window.jQuery && jQuery.fn && jQuery.fn.trigger) {{
            jQuery(lastInput).trigger('change');
        }} else {{
            lastInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }}

        // 强制关闭当前 combo 面板，避免遮挡后续操作
        if (combo && window.jQuery) {{
            try {{ jQuery(combo).combo('hidePanel'); }} catch(e) {{}}
        }}

        return {{
            success: true,
            textInputValue: lastInput.value,
            hiddenInputValue: valueInput ? valueInput.value : null
        }};
    }}
    """
    result = frame.evaluate(js)
    frame.wait_for_timeout(500)
    if result and result.get('success'):
        print(f"[直接写入] {label} = '{text}'")
    else:
        print(f"[警告] 直接写入{label} '{text}' 失败: {result}")
    return result


def select_workitem_tree(frame, workitem_name: str):
    """
    公共工时工作项（combotreegrid 叶子节点选择）：
    1. 关闭其他面板 → 展开下拉加载树数据
    2. 展开所有折叠的父节点，确保目标叶子可见
    3. 用 Playwright 真实 click() 点击匹配的叶子 .tree-node
    4. 验证 input value
    """
    close_all_easyui_panels(frame)
    frame.wait_for_timeout(300)

    # 1. 展开下拉（在公共工时区域内定位）
    _click_combo_arrow_by_placeholder(frame, "工作项", section_hint="公共工时填报")
    frame.wait_for_timeout(1500)

    # 2. 展开所有折叠的父节点
    _expand_all_tree_nodes(frame)

    # 3. 用真实 click 选择叶子节点
    success = _select_from_easyui_panel(frame, workitem_name, "工作项")

    # 4. 验证
    val = get_combo_value_by_placeholder(frame, "工作项")
    if val:
        print(f"[验证] 工作项 input value = '{val}'")
    else:
        print(f"[警告] 工作项 '{workitem_name}' 未选中，请检查截图")

    close_all_easyui_panels(frame)


def _prompt_for_existing_records(date_str: str, page, target_frame, target_day: int) -> str:
    """提示用户处理已有填报记录。返回 'continue' / 'append' / 'stop'"""
    print(f"[警告] {date_str} 检测到已有填报记录！")
    screenshot(page, "02_already_filled")
    try:
        choice = input("请选择操作：1-保留并退出  2-追加填报  3-删除后重新填报 > ").strip()
    except EOFError:
        choice = "1"
    if choice == "2":
        print("[继续] 在现有记录下方追加填报...")
        return "append"
    elif choice == "3":
        print("[提示] 请在系统页面中手动删除所有记录，删除后点击保存草稿，然后回到终端按回车继续...")
        try:
            input("按回车确认已删除并保存...")
        except EOFError:
            pass
        print("[继续] 5秒后重新检查页面状态...")
        page.wait_for_timeout(5000)
        ensure_on_calendar(page)
        target_frame = get_target_frame(page)
        if not target_frame:
            print("[错误] 未找到包含表格的 iframe")
            return "stop"
        if not find_and_click_date(target_frame, target_day):
            print(f"[错误] 重新进入详情页失败")
            return "stop"
        target_frame.wait_for_timeout(1500)
        # 双重检查：既看实际数据，也看是否有输入框
        js_verify = """() => {
            const hasHourInput = document.querySelectorAll("input[type='number']").length > 0;
            const comboValues = Array.from(document.querySelectorAll('.textbox-value')).filter(v => v.value && !isNaN(Number(v.value))).length;
            return { hasHourInput, comboValues };
        }"""
        verify = target_frame.evaluate(js_verify)
        has_draft = is_detail_page_has_draft(target_frame)
        if has_draft and verify.get('hasHourInput') == False and verify.get('comboValues') == 0:
            print("[信息] 详情页已无数据行，判定为已清空（按钮文字误报已排除）。")
            has_draft = False
        if has_draft:
            print("[警告] 重新检查后仍检测到记录，可能未删除干净。脚本终止。")
            return "stop"
        print("[信息] 确认无记录，继续执行填单流程...")
        return "continue"
    else:
        print("[终止] 已保留现有记录，脚本退出。")
        return "stop"


def _get_workitem_value(frame, section_hint: str, index: int = -1):
    js = f"""
    () => {{
        var sections = document.querySelectorAll('div, section');
        var targetSection = null;
        var maxArea = 0;
        for (var i = 0; i < sections.length; i++) {{
            var s = sections[i];
            if ((s.innerText || '').includes('{section_hint}')) {{
                var area = (s.offsetWidth || 0) * (s.offsetHeight || 0);
                if (area > maxArea) {{
                    maxArea = area;
                    targetSection = s;
                }}
            }}
        }}
        var inputs = [];
        if (targetSection) {{
            inputs = targetSection.querySelectorAll("input[placeholder*='工作项']");
        }}
        if (inputs.length === 0) {{
            inputs = document.querySelectorAll("input[placeholder*='工作项']");
        }}
        const idx = {index} < 0 ? inputs.length + ({index}) : {index};
        const target = inputs[idx];
        return target ? (target.value || '') : null;
    }}
    """
    return frame.evaluate(js)


def _direct_fill_combo_text_by_index(frame, keyword: str, text: str, label: str, section_hint: str = None, index: int = -1):
    section_filter = ""
    if section_hint:
        section_filter = f"""
        var sections = document.querySelectorAll('div, section');
        var targetSection = null;
        var maxArea = 0;
        for (var i = 0; i < sections.length; i++) {{
            var s = sections[i];
            if ((s.innerText || '').includes('{section_hint}')) {{
                var area = (s.offsetWidth || 0) * (s.offsetHeight || 0);
                if (area > maxArea) {{
                    maxArea = area;
                    targetSection = s;
                }}
            }}
        }}
        var inputs = [];
        if (targetSection) {{
            inputs = targetSection.querySelectorAll("input[placeholder*='{keyword}']");
        }}
        if (inputs.length === 0) {{
            inputs = document.querySelectorAll("input[placeholder*='{keyword}']");
        }}
        """
    else:
        section_filter = f"""var inputs = document.querySelectorAll("input[placeholder*='{keyword}']");"""

    js = f"""
    () => {{
        {section_filter}
        if (inputs.length === 0) return {{ error: 'no input found' }};
        const idx = {index} < 0 ? inputs.length + ({index}) : {index};
        const lastInput = inputs[idx];
        if (!lastInput) return {{ error: 'index out of range', count: inputs.length, idx: idx }};
        var combo = lastInput.closest('.textbox.combo');
        lastInput.value = '{text}';
        var valueInput = combo ? combo.querySelector('input.textbox-value') : null;
        if (valueInput) {{
            valueInput.value = '{text}';
        }}
        if (window.jQuery && jQuery.fn && jQuery.fn.trigger) {{
            jQuery(lastInput).trigger('change');
        }} else {{
            lastInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }}
        if (combo && window.jQuery) {{
            try {{ jQuery(combo).combo('hidePanel'); }} catch(e) {{}}
        }}
        return {{ success: true, textInputValue: lastInput.value, hiddenInputValue: valueInput ? valueInput.value : null }};
    }}
    """
    result = frame.evaluate(js)
    frame.wait_for_timeout(500)
    if result and result.get('success'):
        print(f"[直接写入] {label} = '{text}'")
    else:
        print(f"[警告] 直接写入{label} '{text}' 失败: {result}")
    return result


def fill_row_inputs(frame, row_index: int, hours: float, log_text: str, is_project: bool):
    """
    填充指定行的工时和日志。
    使用 JS evaluate 绕过 EasyUI 面板遮挡导致的可见性问题。
    """
    hour_inputs = frame.locator("input[type='number']").all()
    textareas = frame.locator("textarea").all()

    js_set_value = """(el, value) => {
        el.value = value;
        el.dispatchEvent(new Event('change', {bubbles: true}));
        el.dispatchEvent(new Event('input', {bubbles: true}));
    }"""

    if row_index < len(hour_inputs):
        hour_inputs[row_index].evaluate(js_set_value, str(hours))
        frame.wait_for_timeout(200)

    if row_index < len(textareas):
        textareas[row_index].evaluate(js_set_value, log_text)
        frame.wait_for_timeout(200)


def _resolve_public_workitem(user_wi: str, db: list[dict]) -> str:
    """将用户描述的公共工作项映射到历史库中存在的标准工作项名称。
    精确匹配 → 模糊匹配 → 默认兜底。
    """
    # 1. 精确匹配历史库中的公共工时工作项3
    for r in db:
        if r.get("申报类型") != "项目工时":
            for k in ("工作项1", "工作项2", "工作项3"):
                if r.get(k) == user_wi:
                    return user_wi
    # 2. 模糊匹配：search_text 中包含用户描述
    matches = []
    for r in db:
        if r.get("申报类型") != "项目工时":
            st = r.get("search_text", "")
            if user_wi.lower() in st.lower():
                matches.append(r)
    if matches:
        # 回退到工作项2/工作项1，避免工作项3为空
        std_wi = matches[0].get("工作项3") or matches[0].get("工作项2") or matches[0].get("工作项1") or user_wi
        print(f"[映射] 将公共工作项 '{user_wi}' 映射到标准工作项 '{std_wi}'")
        return std_wi
    # 3. 兜底
    default_wi = "内部专项工作"
    print(f"[警告] 无法找到公共工作项 '{user_wi}'，使用默认兜底 '{default_wi}'")
    return default_wi


def _match_with_project_coverage(keywords: list[str], db: list[dict], project_log_map: dict | None = None) -> list[dict]:
    """关键词匹配，当存在多项目描述时，确保每个项目关键词都能匹配到记录，避免被单一项目垄断 limit。"""
    if not project_log_map:
        return match_tasks(keywords, db)

    all_matched = []
    seen_ids = set()
    # 1. 为每个项目关键词单独匹配（确保多项目覆盖）
    for proj_kw in project_log_map.keys():
        proj_matched = match_tasks([proj_kw], db, limit=20)
        for r in proj_matched:
            rid = id(r)
            if rid in seen_ids:
                continue
            # 过滤掉与已匹配项目特征重叠的项目工时记录（避免同一项目不同名称重复）
            if r.get("申报类型") == "项目工时":
                r_pname = _safe_str(r.get("项目名称", ""))
                is_dup = any(
                    _has_signature_overlap(r_pname, _safe_str(m.get("项目名称", "")))
                    for m in all_matched if m.get("申报类型") == "项目工时"
                )
                if is_dup:
                    continue
            seen_ids.add(rid)
            all_matched.append(r)
    # 2. 再用所有关键词整体匹配，补充其他相关记录
    overall_matched = match_tasks(keywords, db, limit=30)
    for r in overall_matched:
        rid = id(r)
        if rid in seen_ids:
            continue
        # 过滤掉与已匹配项目特征重叠的项目工时记录（避免同一项目不同名称重复）
        if r.get("申报类型") == "项目工时":
            r_pname = _safe_str(r.get("项目名称", ""))
            is_dup = any(
                _has_signature_overlap(r_pname, _safe_str(m.get("项目名称", "")))
                for m in all_matched if m.get("申报类型") == "项目工时"
            )
            if is_dup:
                continue
        seen_ids.add(rid)
        all_matched.append(r)
    # 按匹配分数重新排序
    all_matched.sort(key=lambda x: x.get("_match_score", 0), reverse=True)
    return all_matched


def _run_fill_impl(date_str: str, keywords: list[str], random_mode: str | None = None, add_tasks: list[str] | None = None, explicit_public: list[dict] | None = None, project_log_map: dict | None = None, project_hour_map: dict | None = None, project_workitem_map: dict | None = None, preview: bool = False):
    # ---- 前置校验 ----
    target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    if target_date > datetime.now().date():
        print(f"[终止] {date_str} 是未来日期，系统不允许提前填报。请选择今天或历史日期。")
        return

    # ---- 数据准备 ----
    db = load_task_db()

    # 加载活跃项目清单（若有）
    active_projects = None
    ap_path = str(ASSETS_DIR / "active_projects.json")
    if os.path.exists(ap_path):
        try:
            from project_index import load_active_projects
            ap_data = load_active_projects(ap_path)
            active_projects = ap_data.get("projects", [])
            last_updated = ap_data.get("last_updated")
            if last_updated:
                from datetime import datetime as _dt
                try:
                    last_dt = _dt.strptime(last_updated, "%Y-%m-%d")
                    days_since = (_dt.now() - last_dt).days
                    if days_since > 30:
                        print(f"[提醒] 距离上次更新活跃项目清单已 {days_since} 天，建议运行 python update_active_projects.py 刷新")
                except Exception:
                    pass
            active_count = sum(1 for p in active_projects if p.get("status") == "active")
            print(f"[信息] 已加载活跃项目清单，active: {active_count} / total: {len(active_projects)}")
        except Exception as e:
            print(f"[警告] 加载活跃项目清单失败: {e}")
    else:
        print("[提醒] 未检测到活跃项目清单（active_projects.json），建议运行 python update_active_projects.py 初始化")

    # 质量策划过滤：完全随机时始终排除；其他模式仅当关键词同时包含"质量策划"+项目名称时才保留
    allow_qp = should_allow_quality_planning(keywords, db)
    if not allow_qp:
        qp_count = sum(1 for r in db if is_quality_planning_related(r))
        if qp_count > 0:
            db = [r for r in db if not is_quality_planning_related(r)]
            print(f"[信息] 已自动排除 {qp_count} 条质量策划记录（随机生成默认不含质量策划）")

    # 预过滤：排除不符合今日日期规则的历史记录，避免匹配/补充到无效内容
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    is_friday = dt.weekday() == 4
    holiday_set = set(config.get("preferences", {}).get("holidays", []))
    next_day = (dt + timedelta(days=1)).strftime("%Y-%m-%d")
    is_day_before_holiday = next_day in holiday_set
    is_late_month = dt.day >= 20
    is_early_month = dt.day <= 10
    eligible_db = []
    filtered_reasons = []
    for r in db:
        leaf = r.get("工作项3", "") or r.get("工作项2", "") or r.get("工作项1", "")
        log = r.get("工作日志", "") or ""
        combined = f"{leaf} {log}"
        if "361评估" in combined and not (is_friday or is_day_before_holiday):
            filtered_reasons.append("361评估")
            continue
        if "审计" in combined and not is_late_month:
            filtered_reasons.append("审计")
            continue
        if "排产" in combined and not is_early_month:
            filtered_reasons.append("排产")
            continue
        eligible_db.append(r)
    if len(eligible_db) < len(db):
        print(f"[信息] 已预过滤 {len(db) - len(eligible_db)} 条不符合今日日期规则的历史记录（{', '.join(sorted(set(filtered_reasons)))}）")
    db = eligible_db

    # === 第1步：生成基础项目工时 ===
    if random_mode == "full":
        total_count = random.choice([6, 7, 8])
        ratios = [(2, 3), (3, 2), (3, 3), (2, 4), (4, 2)]
        proj_cnt, pub_cnt = random.choice(ratios)
        ratio_sum = proj_cnt + pub_cnt
        if ratio_sum != total_count:
            proj_cnt = max(1, round(proj_cnt * total_count / ratio_sum))
            pub_cnt = max(1, total_count - proj_cnt)
        print(f"[随机生成] 总条数 {total_count}，项目 {proj_cnt} / 公共 {pub_cnt}")
        items = weighted_random_tasks(db, total_count, proj_cnt, pub_cnt, date_str, active_projects=active_projects)
    elif random_mode == "supplement":
        matched = _match_with_project_coverage(keywords, db, project_log_map)
        if is_project_specific_keyword(keywords, db):
            matched = [r for r in matched if r.get("申报类型") == "项目工时"]
            if matched:
                print("[信息] 检测到项目特定关键词，已过滤公共工时记录")
        if not matched:
            print("未匹配到历史记录，将完全随机生成")
            total_count = random.choice([6, 7, 8])
            ratios = [(2, 3), (3, 2), (3, 3), (2, 4), (4, 2)]
            proj_cnt, pub_cnt = random.choice(ratios)
            ratio_sum = proj_cnt + pub_cnt
            if ratio_sum != total_count:
                proj_cnt = max(1, round(proj_cnt * total_count / ratio_sum))
                pub_cnt = max(1, total_count - proj_cnt)
            items = weighted_random_tasks(db, total_count, proj_cnt, pub_cnt, date_str, active_projects=active_projects)
        else:
            items = aggregate_tasks(matched)
    else:
        matched = _match_with_project_coverage(keywords, db, project_log_map)
        if is_project_specific_keyword(keywords, db):
            matched = [r for r in matched if r.get("申报类型") == "项目工时"]
            if matched:
                print("[信息] 检测到项目特定关键词，已过滤公共工时记录")
        if not matched:
            if not add_tasks:
                print("未匹配到历史记录，请补充关键词")
                return
            else:
                print("[信息] 未匹配到历史记录，将仅使用手动新增任务")
                items = []
        else:
            items = aggregate_tasks(matched)

    # === 质量策划单项目终身去重 ===
    qp_projects = get_projects_with_quality_planning(load_task_db())
    if qp_projects:
        dedup_logs = []
        for item in items:
            if item.get("申报类型") == "项目工时" and is_quality_planning_related(item):
                pname = item.get("项目名称", "")
                # 用户明确指定的项目保留其质量策划选择
                is_user_specified = any(kw in pname for kw in (project_hour_map or {}).keys())
                if not is_user_specified and pname in qp_projects:
                    old_wi = item["工作项"]
                    item["工作项"] = "项目成本监控"
                    item["工作项1"] = "项目成本监控"
                    item["工作项2"] = ""
                    item["工作项3"] = ""
                    dedup_logs.append(f"[质量策划去重] 项目 '{pname}' 历史上已有质量策划记录，工作项从 '{old_wi}' 改为 '项目成本监控'")
        if dedup_logs:
            print("\n[质量策划终身去重]")
            for log in dedup_logs:
                print(log)

    # 保底项目：用户明确指定但历史库未匹配到的项目，直接构造裸记录，让系统搜索框去匹配
    # 遍历 project_log_map 确保未指定工时的项目也能被兜底（默认1.0H）
    all_user_projects = set()
    if project_log_map:
        all_user_projects.update(project_log_map.keys())
    if project_hour_map:
        all_user_projects.update(project_hour_map.keys())
    if all_user_projects:
        existing_projs = set()
        existing_project_keywords = set()
        # 类别关键词：如果历史库中匹配到了包含子关键词的项目，视为该类别已匹配，不生成保底
        category_keywords = {
            "农业项目": "农业",
            "中移信息": "中移",
        }

        def _strip_suffix(kw: str) -> str:
            for suffix in ["项目", "合同", "服务", "平台", "系统", "实施", "开发", "外包", "运维", "交付", "定制"]:
                if kw.endswith(suffix):
                    return kw[:-len(suffix)]
            return kw

        existing_project_names = set()
        for item in items:
            if item.get("申报类型") == "项目工时":
                pname = item.get("项目名称", "")
                if pname:
                    existing_project_names.add(pname)
                existing_project_keywords.add(_get_project_core_keyword(pname))
                for kw in all_user_projects:
                    if kw in pname:
                        existing_projs.add(kw)
                    # 去掉后缀后宽松匹配，避免"民政局项目"与"深圳民政局"误判为不同项目
                    kw_base = _strip_suffix(kw)
                    if kw_base and kw_base in pname and len(kw_base) >= 2:
                        existing_projs.add(kw)
                    # 类别关键词特殊处理：历史库中已有包含子关键词的项目，视为已匹配
                    if kw in category_keywords and category_keywords[kw] in pname:
                        existing_projs.add(kw)
        for kw in all_user_projects:
            if kw in existing_projs:
                continue
            # 核心关键词去重：若已有同名核心项目，不再追加保底
            kw_core = _get_project_core_keyword(kw)
            if kw_core and kw_core in existing_project_keywords:
                print(f"[信息] 项目 '{kw}' 的核心关键词 '{kw_core}' 已存在匹配记录，跳过保底")
                continue
            # 特征子串重叠检查（如电渠↔电子渠道、H5↔H5）
            if any(_has_signature_overlap(kw, pname) for pname in existing_project_names):
                print(f"[信息] 项目 '{kw}' 的特征子串与已有项目重叠，跳过保底")
                continue
            inferred_wi = (project_workitem_map or {}).get(kw)
            default_wi = inferred_wi if inferred_wi else "项目成本监控"
            w1, w2, w3 = infer_workitem_levels(db, default_wi)
            hours = project_hour_map.get(kw, 1.0) if project_hour_map else 1.0
            items.append({
                "申报类型": "项目工时",
                "项目名称": kw,
                "工作项": w3 or w2 or w1,
                "工作项1": w1,
                "工作项2": w2,
                "工作项3": w3,
                "工时": hours,
                "日志": project_log_map.get(kw, ""),
            })
            print(f"[保底项目] 历史库未匹配到 '{kw}'，将直接使用该关键词去系统搜索填单。工作项: {w3 or w2 or w1} 工时: {hours}H")

    # === 第2步：日期过滤 + 合并项目 ===
    holidays = config.get("preferences", {}).get("holidays", [])
    items, filter_logs = filter_by_date_rules(items, date_str, holidays)
    if filter_logs:
        print("\n[日期规则过滤]")
        for log in filter_logs:
            print(log)

    items = merge_project_rows(items, project_log_map=project_log_map, project_hour_map=project_hour_map)

    # 标记用户指定的项目工时，避免被 adjust_to_eight 缩放
    if project_hour_map:
        for item in items:
            if item.get("申报类型") == "项目工时" and item.get("项目名称"):
                pname = item["项目名称"]
                for kw in project_hour_map.keys():
                    if kw in pname or _has_signature_overlap(kw, pname):
                        item["_user_specified"] = True
                        break

    # === 第3步：添加手动任务 ===
    if add_tasks:
        manual_items = parse_manual_tasks(add_tasks, date_str, db)
        if manual_items:
            print(f"\n[手动任务] 已添加 {len(manual_items)} 条自定义记录")
            for m in manual_items:
                print(f"  - [{m['申报类型']}] {m['项目名称']} | {m['工作项']} | {m['工时']}H")
            items.extend(manual_items)
            # 将手动任务中的项目标记为用户指定，避免被 adjust_to_eight 缩放
            for m in manual_items:
                if m.get("申报类型") == "项目工时" and m.get("项目名称"):
                    project_hour_map[m["项目名称"]] = m["工时"]
                    project_workitem_map[m["项目名称"]] = m["工作项"]

    # === 第4步：添加用户指定的公共工时 ===
    explicit_public = explicit_public or []
    for ep in explicit_public:
        resolved_wi = _resolve_public_workitem(ep["工作项"], db)
        items.append({
            "申报类型": "通用报工",
            "项目名称": None,
            "工作项": resolved_wi,
            "工时": ep["工时"],
            "日志": ep["日志"],
            "业务集团": "ABG",
            "业务线": "电信与AIoT业务线",
            "_user_specified": True,
        })
        print(f"[指定项] 已追加公共工时：{resolved_wi} {ep['工时']}H")

    # === 第5步：添加固定公共工时 ===
    FIXED_PUBLIC_TASK = {
        "申报类型": "通用报工",
        "项目名称": None,
        "工作项": "AI工具引入与推广策划",
        "工时": 2.0,
        "日志": "集团转型AI组织运作与AI工具应用推广相关工作，对接公司布道师、金种子等组织运作及专项工作，培训策划、工作信息下达与传递沟通，各组织优秀实践识别与分享运作等例行工作",
        "业务集团": "ABG",
        "业务线": "电信与AIoT业务线",
        "_fixed": True,
    }
    has_fixed = any(i.get("工作项") == FIXED_PUBLIC_TASK["工作项"] for i in items)
    if not has_fixed:
        items.append(FIXED_PUBLIC_TASK.copy())
        print("[固定项] 已追加每日必出公共工时：AI工具引入与推广策划 2.0H")

    # === 第5.5步：月末例行工作（质量月报 + 审计/度量） ===
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    if dt.day >= 25:
        # 质量月报（公共工时）
        has_month_end_report = any(
            i.get("申报类型") != "项目工时" and "质量月报" in (i.get("日志", "") or "")
            for i in items
        )
        if not has_month_end_report:
            w1, w2, w3 = infer_workitem_levels(db, "部门例行运作")
            items.append({
                "申报类型": "通用报工",
                "项目名称": None,
                "工作项": w3 or w2 or w1,
                "工作项1": w1,
                "工作项2": w2,
                "工作项3": w3,
                "工时": 1.5,
                "日志": "质量月报内容整理刷新，输出质量月报",
                "业务集团": "ABG",
                "业务线": "电信与AIoT业务线",
            })
            print("[月末固定项] 已追加质量月报：部门例行运作 1.5H")

        # 审计/度量（项目工时，按活跃项目随机分摊）
        month_end_tasks = _generate_month_end_tasks(items, active_projects, date_str, db)
        items.extend(month_end_tasks)

    # === 第6步：确保有公共工时 ===
    public_count = sum(1 for i in items if i["申报类型"] != "项目工时")
    if public_count == 0:
        print("[补充随机] 当前无公共工时，强制补充2条")
        public_pool = [r for r in db if r.get("申报类型") != "项目工时"]
        if public_pool:
            extra_public = weighted_random_tasks(public_pool, 2, 0, 2, date_str)
            items.extend(extra_public)

    # === 第7步：条数/工时不足时从历史库补充 ===
    total_hours = sum(i["工时"] for i in items)
    if len(items) < 5 or total_hours < 8.0:
        target = random.choice([6, 7, 8])
        if len(items) < target:
            print(f"[补充随机] 当前 {len(items)} 条/{total_hours}H，补充至 {target} 条")
            items = supplement_with_random(items, db, target, date_str, active_projects=active_projects)
            items, extra_filter_logs = filter_by_date_rules(items, date_str, holidays)
            if extra_filter_logs:
                print("\n[日期规则过滤-补充记录]")
                for log in extra_filter_logs:
                    print(log)
        elif total_hours < 8.0:
            # 条数已够但工时不足，仍补充1条以增加工时
            print(f"[补充随机] 当前 {len(items)} 条/{total_hours}H，补充工时")
            items = supplement_with_random(items, db, len(items) + 1, date_str, active_projects=active_projects)
            items, extra_filter_logs = filter_by_date_rules(items, date_str, holidays)
            if extra_filter_logs:
                print("\n[日期规则过滤-补充记录]")
                for log in extra_filter_logs:
                    print(log)

    # 补充后若总工时超8H，删除或压缩多余补充项
    total_hours = sum(i["工时"] for i in items)
    if total_hours > 8.0:
        excess = total_hours - 8.0
        supplemented_items = [i for i in items if i.get("_supplemented") and i.get("工作项") != "AI工具引入与推广策划"]
        supplemented_items.sort(key=lambda x: x["工时"], reverse=True)
        for item in supplemented_items:
            if excess <= 0:
                break
            if item["工时"] <= excess + 0.5:
                items.remove(item)
                excess -= item["工时"]
                print(f"[工时调整] 删除补充项 '{item['工作项']}' {item['工时']}H，总工时超限")
            else:
                item["工时"] -= excess
                if item["工时"] < 0.5:
                    item["工时"] = 0.5
                excess = 0
                print(f"[工时调整] 压缩补充项 '{item['工作项']}' 至 {item['工时']}H，总工时超限")

    # === 第8步：截断到8条 ===
    if len(items) > 8:
        print(f"[信息] 当前共 {len(items)} 条，截断至8条")
        public_items = [i for i in items if i["申报类型"] != "项目工时"]
        project_items = [i for i in items if i["申报类型"] == "项目工时"]
        # 优先保留固定公共工时 AI工具引入与推广策划
        fixed_public = [i for i in public_items if i.get("工作项") == "AI工具引入与推广策划"]
        other_public = [i for i in public_items if i.get("工作项") != "AI工具引入与推广策划"]
        # 若用户指定了≥2个显式公共，优先保留3个公共（固定+2显式），否则保留2个
        max_public = 3 if len(other_public) >= 2 else 2
        public_items = fixed_public + other_public[:max(0, max_public - len(fixed_public))]
        keep_project = max(1, 8 - len(public_items))
        specified_projects = set(project_log_map.keys()) if project_log_map else set()
        specified_project_items = [i for i in project_items if any(kw in i["项目名称"] for kw in specified_projects)]
        other_project_items = [i for i in project_items if not any(kw in i["项目名称"] for kw in specified_projects)]
        other_project_items.sort(key=lambda x: x["工时"], reverse=True)
        remaining_slots = max(0, keep_project - len(specified_project_items))
        project_items = specified_project_items[:keep_project] + other_project_items[:remaining_slots]
        items = project_items + public_items

    # 确保 project_items 和 public_items 始终定义（用于后续填报循环）
    project_items = [i for i in items if i["申报类型"] == "项目工时"]
    public_items = [i for i in items if i["申报类型"] != "项目工时"]

    # === 第9步：月份替换 ===
    month_logs = replace_month_in_logs(items, date_str)
    if month_logs:
        print("\n[月份关联性替换]")
        for log in month_logs:
            print(log)

    # 应用用户明确指定的项目工作内容，覆盖对应项目的日志和工作项（二次保险）
    project_log_map = project_log_map or {}
    if project_log_map:
        for item in items:
            if item.get("申报类型") == "项目工时" and item.get("项目名称"):
                pname = item["项目名称"]
                for proj_kw, log_content in project_log_map.items():
                    if proj_kw in pname:
                        item["日志"] = log_content
                        print(f"[覆盖日志] 项目 '{pname}' 的日志已替换为指定内容")
                        break

    project_workitem_map = project_workitem_map or {}
    if project_workitem_map:
        for item in items:
            if item.get("申报类型") == "项目工时" and item.get("项目名称"):
                pname = item["项目名称"]
                for proj_kw, wi_content in project_workitem_map.items():
                    if proj_kw in pname:
                        old_wi = item.get("工作项", "")
                        item["工作项"] = wi_content
                        item["工作项1"] = wi_content
                        item["工作项2"] = ""
                        item["工作项3"] = ""
                        if old_wi != wi_content:
                            print(f"[覆盖工作项] 项目 '{pname}' 的工作项从 '{old_wi}' 替换为 '{wi_content}'")
                        break

    # === 第9.5步：合并重复公共工时 ===
    # 对通用报工（非项目工时）按工作项合并，避免同一工作项出现多行
    # 固定公共工时 AI工具引入与推广策划（2.0H）不参与合并，仅保留固定值，丢弃同名补充项
    public_map = {}
    merged_items = []
    FIXED_WI = "AI工具引入与推广策划"
    for item in items:
        if item.get("申报类型") != "项目工时":
            wi = item.get("工作项", "")
            if wi == FIXED_WI:
                # 固定工时只保留第一条（即固定项2.0H），丢弃后续同名的随机补充项
                if FIXED_WI not in public_map:
                    public_map[wi] = item.copy()
                else:
                    print(f"[固定项保护] 丢弃同名补充项 '{wi}' {item['工时']}H，固定项保持 {public_map[wi]['工时']}H")
                continue
            if wi in public_map:
                public_map[wi]["工时"] += item["工时"]
                existing_log = public_map[wi].get("日志", "")
                new_log = item.get("日志", "")
                if new_log and new_log not in existing_log:
                    public_map[wi]["日志"] = existing_log + "；" + new_log if existing_log else new_log
                    if len(public_map[wi]["日志"]) > 500:
                        public_map[wi]["日志"] = public_map[wi]["日志"][:500]
                print(f"[公共工时合并] 工作项 '{wi}' 工时合并为 {public_map[wi]['工时']}H")
            else:
                public_map[wi] = item.copy()
        else:
            merged_items.append(item)
    merged_items.extend(public_map.values())
    items = merged_items

    # === 第10步：统一调整工时 ===
    # 公共工时 + 用户明确指定的项目工时 = 固定预留，不参与缩放
    # 仅对非用户指定的项目工时进行缩放
    dynamic_items = [i for i in items if i.get("申报类型") == "项目工时"]
    public_items = [i for i in items if i.get("申报类型") != "项目工时"]
    user_specified_items = []
    other_dynamic_items = []
    project_hour_map = project_hour_map or {}
    for item in dynamic_items:
        pname = item.get("项目名称", "")
        is_specified = False
        for kw in project_hour_map.keys():
            if kw in pname:
                item["工时"] = project_hour_map[kw]
                is_specified = True
                break
        if is_specified:
            user_specified_items.append(item)
        else:
            other_dynamic_items.append(item)

    public_hours = sum(i["工时"] for i in public_items)
    reserved_hours = public_hours + sum(i["工时"] for i in user_specified_items)
    if other_dynamic_items:
        other_dynamic_items = adjust_to_eight(other_dynamic_items, reserved_hours=reserved_hours)
    elif reserved_hours > 8.0:
        # 公共工时（保底+补充）已超8H，压缩补充记录
        supplemented_items = [i for i in items if i.get("_supplemented") and i.get("工作项") != "AI工具引入与推广策划"]
        if supplemented_items:
            excess = reserved_hours - 8.0
            for item in supplemented_items:
                if excess <= 0:
                    break
                reduce = min(item["工时"], excess)
                item["工时"] -= reduce
                item["工时"] = round_to_half(item["工时"])
                excess -= reduce
            print(f"[工时调整] 公共工时超出 8H，已压缩 {len(supplemented_items)} 条补充记录")
        else:
            print(f"[警告] 固定部分已达 {reserved_hours}H，且无补充记录可压缩")
    elif reserved_hours < 8.0:
        # 方案一：无可调项目且总工时不足8H，仅对补充记录进行缩放，不动用户指定部分
        supplemented_items = [i for i in items if i.get("_supplemented") and i.get("工作项") != "AI工具引入与推广策划"]
        if supplemented_items:
            fixed_items = [i for i in items if not i.get("_supplemented")]
            fixed_hours = sum(i["工时"] for i in fixed_items)
            print(f"[工时调整] 固定部分 {fixed_hours}H，对 {len(supplemented_items)} 条补充记录进行缩放以凑齐 8H")
            supplemented_items = adjust_to_eight(supplemented_items, reserved_hours=fixed_hours)
            # adjust_to_eight 已就地修改了字典对象，无需重新组装 items
        else:
            print(f"[警告] 固定部分仅 {reserved_hours}H，且无补充记录可缩放，无法凑齐 8H")
    items = user_specified_items + other_dynamic_items + public_items

    # === 工时覆盖后二次补充 ===
    total_hours = sum(i["工时"] for i in items)
    while total_hours < 8.0 and len(items) < 8:
        print(f"[补充随机] 工时覆盖后当前 {len(items)} 条/{total_hours}H，补充工时")
        items = supplement_with_random(items, db, len(items) + 1, date_str, active_projects=active_projects)
        items, extra_filter_logs = filter_by_date_rules(items, date_str, holidays)
        if extra_filter_logs:
            print("\n[日期规则过滤-补充记录]")
            for log in extra_filter_logs:
                print(log)
        # 若补充后超过8条，再次截断
        if len(items) > 8:
            print(f"[信息] 补充后当前共 {len(items)} 条，截断至8条")
            public_items = [i for i in items if i["申报类型"] != "项目工时"]
            project_items = [i for i in items if i["申报类型"] == "项目工时"]
            fixed_public = [i for i in public_items if i.get("工作项") == "AI工具引入与推广策划"]
            other_public = [i for i in public_items if i.get("工作项") != "AI工具引入与推广策划"]
            max_public = 3 if len(other_public) >= 2 else 2
            public_items = fixed_public + other_public[:max(0, max_public - len(fixed_public))]
            keep_project = max(1, 8 - len(public_items))
            specified_projects = set(project_log_map.keys()) if project_log_map else set()
            specified_project_items = [i for i in project_items if any(kw in i["项目名称"] for kw in specified_projects)]
            other_project_items = [i for i in project_items if not any(kw in i["项目名称"] for kw in specified_projects)]
            other_project_items.sort(key=lambda x: x["工时"], reverse=True)
            remaining_slots = max(0, keep_project - len(specified_project_items))
            project_items = specified_project_items[:keep_project] + other_project_items[:remaining_slots]
            items = project_items + public_items
        total_hours = sum(i["工时"] for i in items)
        if len(items) > 8:
            print(f"[信息] 补充后当前共 {len(items)} 条，截断至8条")
            public_items = [i for i in items if i["申报类型"] != "项目工时"]
            project_items = [i for i in items if i["申报类型"] == "项目工时"]
            fixed_public = [i for i in public_items if i.get("工作项") == "AI工具引入与推广策划"]
            other_public = [i for i in public_items if i.get("工作项") != "AI工具引入与推广策划"]
            max_public = 3 if len(other_public) >= 2 else 2
            public_items = fixed_public + other_public[:max(0, max_public - len(fixed_public))]
            keep_project = max(1, 8 - len(public_items))
            specified_projects = set(project_log_map.keys()) if project_log_map else set()
            specified_project_items = [i for i in project_items if any(kw in i["项目名称"] for kw in specified_projects)]
            other_project_items = [i for i in project_items if not any(kw in i["项目名称"] for kw in specified_projects)]
            other_project_items.sort(key=lambda x: x["工时"], reverse=True)
            remaining_slots = max(0, keep_project - len(specified_project_items))
            project_items = specified_project_items[:keep_project] + other_project_items[:remaining_slots]
            items = project_items + public_items

    # === 单日质量策划项目数硬性约束 ===
    items = enforce_daily_quality_planning_limit(items, limit=1)

    # === 最终强制对齐8H ===
    total = sum(i["工时"] for i in items)
    if total > 8.0:
        excess = total - 8.0
        print(f"[警告] 草稿总工时 {total}H 超过 8H，强制压缩")
        # 优先级1：压缩非固定公共工时（保留 AI工具引入与推广策划）
        for item in sorted([i for i in items if i.get("申报类型") != "项目工时" and i.get("工作项") != "AI工具引入与推广策划"], key=lambda x: x["工时"], reverse=True):
            if excess <= 0:
                break
            reduce = min(item["工时"] - 0.5, excess)
            if reduce > 0:
                item["工时"] -= reduce
                item["工时"] = round_to_half(item["工时"])
                excess -= reduce
                print(f"[最终压缩] '{item['工作项']}' 压缩至 {item['工时']}H")
        # 优先级2：压缩项目工时
        total = sum(i["工时"] for i in items)
        if total > 8.0:
            excess = total - 8.0
            for item in sorted([i for i in items if i.get("申报类型") == "项目工时"], key=lambda x: x["工时"], reverse=True):
                if excess <= 0:
                    break
                reduce = min(item["工时"] - 0.5, excess)
                if reduce > 0:
                    item["工时"] -= reduce
                    item["工时"] = round_to_half(item["工时"])
                    excess -= reduce
                    print(f"[最终压缩] '{item['工作项']}' 压缩至 {item['工时']}H")
        # 优先级3：压缩固定公共工时（最后手段）
        total = sum(i["工时"] for i in items)
        if total > 8.0:
            excess = total - 8.0
            for item in sorted([i for i in items if i.get("工作项") == "AI工具引入与推广策划"], key=lambda x: x["工时"], reverse=True):
                if excess <= 0:
                    break
                reduce = min(item["工时"] - 0.5, excess)
                if reduce > 0:
                    item["工时"] -= reduce
                    item["工时"] = round_to_half(item["工时"])
                    excess -= reduce
                    print(f"[最终压缩] '{item['工作项']}' 压缩至 {item['工时']}H")
        # 极端情况：删除最小工时项
        total = sum(i["工时"] for i in items)
        while total > 8.0:
            candidates = [i for i in items if i.get("工作项") != "AI工具引入与推广策划"]
            if not candidates:
                candidates = items
            if not candidates:
                break
            to_remove = min(candidates, key=lambda x: x["工时"])
            items.remove(to_remove)
            print(f"[最终删除] 删除 '{to_remove['工作项']}' {to_remove['工时']}H，总工时超限")
            total = sum(i["工时"] for i in items)

    if not items:
        print("\n[终止] 日期规则过滤后无可用记录，无需填报。")
        return

    total = sum(i["工时"] for i in items)

    print(f"\n=== 工作日志草稿（{date_str}）===\n")
    # 按表格样式输出，便于用户对照系统页面调整
    print(f"{'序号':<4} | {'类型':<4} | {'项目':<30} | {'工作项':<20} | {'工时':<5} | {'详细日志'}")
    print("-" * 100)
    for idx, item in enumerate(items, 1):
        log = item.get('日志', '')
        if len(log) > 500:
            log = log[:500] + "...（已截断至500字）"
        itype = "项目" if item["申报类型"] == "项目工时" else "公共"
        proj = item.get('项目名称', '') or '-'
        proj_disp = proj[:28] + "..." if len(proj) > 30 else proj
        wi = item.get('工作项', '')
        wi_disp = wi[:18] + "..." if len(wi) > 20 else wi
        inactive_tag = " [历史库/非活跃]" if item.get("_from_inactive") else ""
        print(f"{idx:<4} | {itype:<4} | {proj_disp:<30} | {wi_disp:<20} | {item['工时']:<5}H | {log}{inactive_tag}")
    print("-" * 100)
    print(f"合计: {total}H\n")
    print("如需调整请修改关键词后重新运行。\n")

    if preview:
        print("[预览模式] 草稿已生成，未连接浏览器。请确认内容无误后再执行正式填单。\n")
        return

    # ---- 浏览器自动化 ----
    with sync_playwright() as p:
        print(f"正在连接 Chrome ({CDP_URL})...")
        try:
            browser = p.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            print(f"连接失败: {e}")

            # 尝试自动启动 Chrome
            launcher_path = os.path.join(os.path.dirname(__file__), "launcher.ps1")
            if os.path.exists(launcher_path):
                print(f"[信息] 尝试自动启动 Chrome: {launcher_path}")
                result = subprocess.run(
                    ["powershell", "-ExecutionPolicy", "Bypass", "-File", launcher_path],
                    capture_output=True, text=True
                )
                print(result.stdout)
                if result.returncode != 0:
                    print(result.stderr)
                    print("[错误] 自动启动 Chrome 失败，请手动启动")
                    return

                # 等待 Chrome 调试端口就绪
                print("[信息] 等待 Chrome 调试端口就绪...")
                for i in range(30):
                    try:
                        urllib.request.urlopen("http://127.0.0.1:9222/json/version", timeout=1)
                        print("[成功] Chrome 调试端口已就绪")
                        break
                    except Exception:
                        time.sleep(1)
                else:
                    print("[错误] 等待 Chrome 启动超时（30秒），请检查")
                    return

                # 重试连接
                print("正在重试连接 Chrome...")
                try:
                    browser = p.chromium.connect_over_cdp(CDP_URL)
                except Exception as e2:
                    print(f"重试连接失败: {e2}")
                    return
            else:
                print("请确认 Chrome 已使用 --remote-debugging-port=9222 启动")
                return

        # 遍历所有页面，找到目标系统页面
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                print(f"发现页面: {pg.url}")
                if "prj.chinasoftinc.com" in pg.url:
                    page = pg
                    break
            if page:
                break
        if not page:
            print("未找到目标系统页面，请确认已登录并打开系统页面")
            return
        print(f"已选中页面: {page.url}")

        ensure_on_calendar(page)
        screenshot(page, "01_calendar")

        target_frame = get_target_frame(page)
        if not target_frame:
            print("[错误] 未找到包含表格的 iframe")
            return

        target_day = datetime.strptime(date_str, "%Y-%m-%d").day

        has_draft = is_date_filled(target_frame, target_day)
        action = None

        if has_draft:
            action = _prompt_for_existing_records(date_str, page, target_frame, target_day)
            if action == "stop":
                return

        if not find_and_click_date(target_frame, target_day):
            print(f"[错误] 未找到日期 {target_day} 的单元格")
            screenshot(page, "02_date_not_found")
            return

        target_frame.wait_for_timeout(1500)

        # 日历检测漏检时，详情页二次检测
        if not has_draft and is_detail_page_has_draft(target_frame):
            action = _prompt_for_existing_records(date_str, page, target_frame, target_day)
            if action == "stop":
                return

        screenshot(page, "03_detail_page")

        # ---- 填报项目工时 ----
        failed_workitems = []
        filled_project_names = {}  # idx -> 系统回填的完整项目名称
        for idx, item in enumerate(project_items):
            add_row_in_section(target_frame, "项目工时填报")
            filled_name = select_project(target_frame, item["项目名称"])
            if filled_name:
                filled_project_names[idx] = filled_name
            ok = select_workitem_dropdown(target_frame, item["工作项"], row_index=idx)
            if not ok:
                failed_workitems.append((item["项目名称"], item["工作项"]))
                continue
            fill_row_inputs(target_frame, idx, item["工时"], item["日志"], is_project=True)
            screenshot(page, f"04_project_row_{idx}")

        if failed_workitems:
            print(f"\n[警告] 以下 {len(failed_workitems)} 条项目工时工作项选择失败，已跳过工时和日志填写：")
            for proj, wi in failed_workitems:
                print(f"  - {proj} | {wi}")
            print("[提示] 请在系统页面中手动补充这些行的工作项、工时和日志，或删除空行后保存草稿。\n")

        # ---- 填报公共工时 ----
        public_items = public_items
        for idx, item in enumerate(public_items):
            add_row_in_section(target_frame, "公共工时填报")
            select_department(target_frame, default_department)
            select_workitem_tree(target_frame, item["工作项"])
            row_idx = len(project_items) + idx
            fill_row_inputs(target_frame, row_idx, item["工时"], item["日志"], is_project=False)
            screenshot(page, f"05_public_row_{idx}")

        # ---- 保存草稿 ----
        close_all_easyui_panels(target_frame)
        try:
            target_frame.get_by_text("保存草稿").click()
        except Exception as e:
            print(f"[警告] Playwright 点击保存草稿失败: {e}，尝试 JS click...")
            target_frame.evaluate("""() => { var btn = document.getElementById('btnSaveDraft'); if (btn) btn.click(); }""")
        target_frame.wait_for_timeout(1500)
        screenshot(page, "06_draft_saved")
        print("[成功] 草稿已保存。请务必人工核对系统页面内容无误后，再手动点击【提交】按钮。脚本不会自动提交。")

        # 持久化手动新增任务到任务库
        if add_tasks:
            manual_items = parse_manual_tasks(add_tasks, date_str, db)
            if manual_items:
                append_to_task_db(manual_items, date_str)

        # 自动持久化本次填单记录到任务库（去重）
        try:
            records_to_append = []
            for idx, item in enumerate(project_items):
                full_name = filled_project_names.get(idx, item["项目名称"])
                records_to_append.append({
                    "申报类型": "项目工时",
                    "项目名称": full_name,
                    "工作项": item.get("工作项", ""),
                    "工作项1": item.get("工作项1", item.get("工作项", "")),
                    "工作项2": item.get("工作项2", ""),
                    "工作项3": item.get("工作项3", ""),
                    "工时": item["工时"],
                    "日志": item["日志"],
                })
            for item in public_items:
                records_to_append.append({
                    "申报类型": "通用报工",
                    "项目名称": None,
                    "工作项": item.get("工作项", ""),
                    "工作项1": item.get("工作项1", item.get("工作项", "")),
                    "工作项2": item.get("工作项2", ""),
                    "工作项3": item.get("工作项3", ""),
                    "工时": item["工时"],
                    "日志": item["日志"],
                })

            with open(str(ASSETS_DIR / "task_db.json"), "r", encoding="utf-8") as f:
                existing_db = json.load(f)
            existing_keys = set()
            for r in existing_db:
                key = (r.get("项目名称") or "", r.get("工作日志") or "", r.get("工作项1") or r.get("工作项3") or r.get("工作项2") or "")
                existing_keys.add(key)

            unique_records = []
            for rec in records_to_append:
                key = (rec.get("项目名称") or "", rec.get("日志") or "", rec.get("工作项1") or "")
                if key not in existing_keys:
                    unique_records.append(rec)
                else:
                    print(f"[信息] 记录已存在于历史库，跳过追加: {rec.get('项目名称')} | {rec.get('工作项1')}")

            if unique_records:
                append_to_task_db(unique_records, date_str)
        except Exception as e:
            print(f"[警告] 自动持久化到任务库失败: {e}")


def parse_daily_description(text: str, active_projects: list[dict] | None = None) -> dict:
    """解析自然语言描述，提取日期、关键词和明确指定的公共任务。
    返回 {"date": "YYYY-MM-DD", "keywords": [...], "explicit_public": [{"工作项": ..., "工时": ..., "日志": ...}]}
    """
    import re
    from datetime import datetime

    result = {"date": None, "keywords": [], "explicit_public": [], "project_log_map": {}, "project_hour_map": {}, "project_workitem_map": {}}
    now = datetime.now()

    # 日期提取（仅当文本中包含明确的日期表达时才提取）
    if "今天" in text:
        result["date"] = now.strftime("%Y-%m-%d")
    elif "昨天" in text:
        from datetime import timedelta
        result["date"] = (now - timedelta(days=1)).strftime("%Y-%m-%d")
    elif "大前天" in text:
        from datetime import timedelta
        result["date"] = (now - timedelta(days=3)).strftime("%Y-%m-%d")
    elif "前天" in text:
        from datetime import timedelta
        result["date"] = (now - timedelta(days=2)).strftime("%Y-%m-%d")
    else:
        # 匹配 M.D 或 M月D日，排除小数点（如 0.5小时）和明显不合理的日期
        m = re.search(r"(?<!\d)([1-9]|1[0-2])\.([1-9]|[12]\d|3[01])(?!\d)", text)
        if not m:
            m = re.search(r"([1-9]|1[0-2])月([1-9]|[12]\d|3[01])日", text)
        if m:
            month, day = int(m.group(1)), int(m.group(2))
            result["date"] = f"{now.year}-{month:02d}-{day:02d}"

    def _clean_kw(s: str) -> str:
        """去除关键词末尾的工时、标点等多余信息。"""
        s = re.sub(r"[0-9.]+\s*(?:h|H|小时)\s*$", "", s)
        s = s.strip("，,、；;。. ")
        return s

    def _simplify_project_kw(s: str) -> list[str]:
        """对项目关键词进行简化，生成多个备选。"""
        variants = [s]
        # 只有当输入中不包含明确年份时，才生成无年份变体
        # 避免用户输入"2026年XX项目"时同时匹配到不带年份版本的历史记录
        if not re.search(r"\d{4}", s):
            # 去除年份如 2028年、2026-2028年
            no_year = re.sub(r"\d{4}[-~]\d{4}年?", "", s)
            no_year = re.sub(r"\d{4}年?", "", no_year)
            no_year = no_year.strip(" -")
            if no_year and no_year != s and len(no_year) >= 2:
                variants.append(no_year)
        # 去除"项目"后缀
        no_proj = re.sub(r"项目$", "", s).strip()
        if no_proj and no_proj != s and len(no_proj) >= 2:
            variants.append(no_proj)
        return variants

    # 提取"项目级"部分
    proj_match = re.search(r"项目(?:级|工时)[：:](.+?)(?:公共(?:工时)?[：:]|$)", text, re.DOTALL)
    if proj_match:
        proj_section = proj_match.group(1).strip()
        raw_projects = re.split(r"[；;]", proj_section)
        projects = []
        for rp in raw_projects:
            rp = rp.strip()
            if not rp:
                continue
            hour_matches = list(re.finditer(r"[0-9.]+\s*(?:h|H|小时)", rp))
            if len(hour_matches) > 1:
                start = 0
                for m in hour_matches:
                    end = m.end()
                    projects.append(rp[start:end].strip("，,；; "))
                    start = end
                    while start < len(rp) and rp[start] in "，,；; ":
                        start += 1
            else:
                projects.append(rp)
        for p in projects:
            p = p.strip()
            if not p:
                continue
            parts = p.split("，", 1)
            if len(parts) >= 1:
                kw = _clean_kw(parts[0].strip())
                if kw:
                    for variant in _simplify_project_kw(kw):
                        result["keywords"].append(variant)
                        # 对地名尝试简化（如深圳市民政局 → 深圳民政局）
                        simple = re.sub(r"([省市县区])\1", r"\1", variant)  # 去重
                        simple = simple.replace("市", "")
                        if simple != variant and len(simple) >= 2:
                            result["keywords"].append(simple)
            proj_key = _clean_kw(parts[0].strip()) if parts else ""
            if proj_key:
                # 生成项目关键词变体（含地名简化），用于匹配和日志/工时覆盖
                proj_variants = set([proj_key])
                for variant in _simplify_project_kw(proj_key):
                    proj_variants.add(variant)
                    simple = re.sub(r"([省市县区])\1", r"\1", variant)
                    simple = simple.replace("市", "")
                    if simple != variant and len(simple) >= 2:
                        proj_variants.add(simple)
                for v in proj_variants:
                    result["keywords"].append(v)

                # 尝试将项目简称映射到完整名称（基于活跃项目清单）
                if active_projects is not None:
                    from project_index import resolve_project_name
                    resolved = resolve_project_name(proj_key, {"projects": active_projects})
                    if resolved and resolved not in proj_variants:
                        proj_variants.add(resolved)
                        result["keywords"].append(resolved)

            if len(parts) >= 2:
                work_kw = _clean_kw(parts[1].strip())
                if work_kw:
                    result["keywords"].append(work_kw)
                    # 记录项目→工作内容映射，用于后续覆盖日志
                    if proj_key:
                        for v in proj_variants:
                            result["project_log_map"][v] = work_kw
                    # 根据工作内容关键词推断工作项
                    inferred_wi = _infer_workitem_from_log(work_kw)
                    if inferred_wi and proj_key:
                        for v in proj_variants:
                            result["project_workitem_map"][v] = inferred_wi
            # 提取项目明确指定的工时
            hour_match = re.search(r"([0-9.]+)\s*(?:h|H|小时)", p)
            if hour_match and proj_key:
                for v in proj_variants:
                    result["project_hour_map"][v] = float(hour_match.group(1))

    # 提取"公共"部分
    public_match = re.search(r"公共(?:工时)?[：:](.+?)$", text, re.DOTALL)
    if public_match:
        public_section = public_match.group(1).strip()
        public_section = re.sub(r"^工时\s*[：:]?", "", public_section).strip()
        public_items = re.split(r"[；;]", public_section)
        for item in public_items:
            item = item.strip()
            if not item:
                continue
            hour_match = re.search(r"([0-9.]+)\s*(?:h|H|小时)", item)
            hours = float(hour_match.group(1)) if hour_match else None
            work_desc = _clean_kw(re.sub(r"[0-9.]+\s*(?:h|H|小时)[,，]?", "", item))
            if work_desc:
                result["keywords"].append(work_desc)
                if hours is not None:
                    result["explicit_public"].append({
                        "工作项": work_desc,
                        "工时": hours,
                        "日志": work_desc,
                    })
    return result


def run_fill(date_str: str, keywords: list[str], random_mode: str | None = None, add_tasks: list[str] | None = None, explicit_public: list[dict] | None = None, project_log_map: dict | None = None, project_hour_map: dict | None = None, project_workitem_map: dict | None = None, preview: bool = False):
    """包装 _run_fill_impl，自动将控制台输出同时写入 diagnose_report.md。"""
    import datetime as _dt
    import os

    class TeeLogger:
        def __init__(self, filename, stream):
            self.file = open(filename, "a", encoding="utf-8")
            self.stream = stream
            header = f"\n\n---\n**执行时间**: {_dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            self.file.write(header)
            self.file.flush()

        def write(self, data):
            self.file.write(data)
            self.stream.write(data)
            self.file.flush()

        def flush(self):
            self.file.flush()
            self.stream.flush()

        def close(self):
            self.file.close()

    log_path = os.path.join(os.path.dirname(__file__), "diagnose_report.md")
    tee = TeeLogger(log_path, sys.stdout)
    original_stdout = sys.stdout
    sys.stdout = tee
    try:
        _run_fill_impl(date_str, keywords, random_mode=random_mode, add_tasks=add_tasks, explicit_public=explicit_public, project_log_map=project_log_map, project_hour_map=project_hour_map, project_workitem_map=project_workitem_map, preview=preview)
    finally:
        sys.stdout = original_stdout
        tee.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="自动填报工作日志")
    parser.add_argument("--date", help="目标日期 YYYY-MM-DD，支持 today/yesterday/before_yesterday/3days_ago（默认今天）")
    parser.add_argument("--keywords", help="关键词，空格分隔")
    parser.add_argument("--random", choices=["full", "supplement"], help="随机生成模式: full=完全随机, supplement=关键词匹配后补充随机")
    parser.add_argument("--add-task", action="append", help="手动新增任务，格式: 申报类型|项目名称|工作项|工时|日志，可多次传入")
    parser.add_argument("--describe", help="自然语言描述今日工作，自动提取日期、关键词和明确工时")
    parser.add_argument("--preview", action="store_true", help="仅生成草稿预览，不连接浏览器填单")
    args = parser.parse_args()

    explicit_public = []
    project_log_map = {}
    project_hour_map = {}
    project_workitem_map = {}

    # 加载活跃项目清单（用于 parse_daily_description 的项目简称映射）
    active_projects_for_parse = None
    ap_path = str(ASSETS_DIR / "active_projects.json")
    if os.path.exists(ap_path):
        try:
            from project_index import load_active_projects
            ap_data = load_active_projects(ap_path)
            active_projects_for_parse = ap_data.get("projects", [])
        except Exception:
            pass

    if args.describe:
        parsed = parse_daily_description(args.describe, active_projects=active_projects_for_parse)
        # --describe 解析出的日期仅作为后备，不覆盖用户显式传入的 --date
        if parsed["date"] and not args.date:
            args.date = parsed["date"]
        if parsed["keywords"]:
            args.keywords = " ".join(parsed["keywords"])
        explicit_public = parsed["explicit_public"]
        project_log_map = parsed.get("project_log_map", {})
        project_hour_map = parsed.get("project_hour_map", {})
        project_workitem_map = parsed.get("project_workitem_map", {})

    # 处理日期：支持自然语言或省略时默认今天
    from datetime import datetime, timedelta
    if args.date:
        dl = args.date.lower()
        if dl == "today":
            args.date = datetime.now().strftime("%Y-%m-%d")
        elif dl == "yesterday":
            args.date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        elif dl in ("before_yesterday", "before-yesterday", "前天"):
            args.date = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")
        elif dl in ("3days_ago", "3-days-ago", "大前天"):
            args.date = (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d")
    if not args.date:
        args.date = datetime.now().strftime("%Y-%m-%d")

    if not args.random and not args.keywords and not args.add_task:
        parser.error("非随机模式下必须提供 --keywords 或 --add-task")
    if args.random == "supplement" and not args.keywords and not args.add_task:
        parser.error("--random supplement 模式下必须提供 --keywords 或 --add-task")

    keywords = args.keywords.split() if args.keywords else []
    run_fill(args.date, keywords, random_mode=args.random, add_tasks=args.add_task, explicit_public=explicit_public, project_log_map=project_log_map, project_hour_map=project_hour_map, project_workitem_map=project_workitem_map, preview=args.preview)
