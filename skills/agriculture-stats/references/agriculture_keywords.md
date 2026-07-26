# 农业项目识别关键字

**用途**: 用于从项目名称中自动识别和筛选农业相关项目  
**更新时间**: 2026-03-20

---

## 关键字列表

以下关键字用于匹配农业行业项目（不区分大小写，包含即匹配）：

```
农业科学院
农科院
农业农村部
农村
农田
种植
养殖
畜牧
渔业
农机
种业
农资
农产品
畜禽
林业
农场
全国畜牧总站
审计署
兽医
中监所
```

---

## 排除关键字（可选）

以下关键字出现时，即使包含上述农业关键字，也应排除：

```
农业项目群
银行
保险
证券
贷款
融资
```

---

## 使用方法

### Python 代码示例

```python
from pathlib import Path

def load_agriculture_keywords(file_path: str = None):
    """加载农业项目关键字"""
    if file_path is None:
        file_path = Path(__file__).parent / 'agriculture_keywords.md'

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 解析关键字列表
    include_keywords = []
    exclude_keywords = []

    in_include_section = False
    in_exclude_section = False

    for line in content.split('\n'):
        line = line.strip()

        if line == '## 关键字列表':
            in_include_section = True
            in_exclude_section = False
            continue
        elif line == '## 排除关键字（可选）':
            in_include_section = False
            in_exclude_section = True
            continue
        elif line.startswith('##'):
            in_include_section = False
            in_exclude_section = False
            continue

        # 跳过代码块标记和空行
        if line.startswith('```') or not line:
            continue

        # 添加关键字
        if in_include_section:
            include_keywords.append(line)
        elif in_exclude_section:
            exclude_keywords.append(line)

    return include_keywords, exclude_keywords


def is_agriculture_project(project_name: str, include_keywords: list, exclude_keywords: list) -> bool:
    """
    判断项目是否为农业项目

    Args:
        project_name: 项目名称
        include_keywords: 包含关键字列表
        exclude_keywords: 排除关键字列表

    Returns:
        bool: True 表示是农业项目
    """
    # 先检查排除关键字
    for keyword in exclude_keywords:
        if keyword in project_name:
            return False

    # 再检查农业关键字
    for keyword in include_keywords:
        if keyword in project_name:
            return True

    return False


# 使用示例
if __name__ == '__main__':
    include_kw, exclude_kw = load_agriculture_keywords()

    test_projects = [
        '河南 -2025 年农业农村监测平台建设项目-FP 项目',
        '某银行农业贷款服务项目',
        '智慧农业大数据平台',
        '某某公司 IT 系统升级项目',
        '乡村振兴示范项目',
    ]

    print("农业项目识别测试:\n")
    for project in test_projects:
        is_agri = is_agriculture_project(project, include_kw, exclude_kw)
        status = "✓" if is_agri else "✗"
        print(f"{status} {project}")
```

---

## 维护说明

- **添加关键字**: 直接在对应列表中新增一行
- **删除关键字**: 删除对应行即可
- **关键字格式**: 每行一个关键字，不需要引号或逗号

---

**最后更新**: 2026-03-20
