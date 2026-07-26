import json
import sys
import os
import html
import argparse


def escape(text):
    return html.escape(str(text))


def build_conclusion_items(items):
    lines = []
    for item in items:
        lines.append(f'                <li>{escape(item)}</li>')
    return "\n".join(lines)


def build_card(card):
    color = card.get("color", "blue")
    items_html = "\n".join(
        f'                        <li>{escape(item)}</li>' for item in card.get("items", [])
    )
    return f'''            <!-- {escape(card.get("title", ""))} -->
            <div class="card card-{color}">
                <div class="card-header">
                    <span class="emoji">{escape(card.get("emoji", ""))}</span>
                    <h3>{escape(card.get("title", ""))}</h3>
                </div>
                <div class="card-inner">
                    <h4>{escape(card.get("subtitle", ""))}</h4>
                    <ul>
{items_html}
                    </ul>
                </div>
            </div>'''


def build_cards(cards):
    return "\n".join(build_card(c) for c in cards)


def build_action_rows(actions):
    lines = []
    last_idx = len(actions) - 1
    for idx, action in enumerate(actions):
        border_class = "" if idx == last_idx else " border-b"
        lines.append(
            f'                            <tr>'
        )
        lines.append(f'                                <td>{escape(action.get("no", ""))}</td>')
        lines.append(f'                                <td>{escape(action.get("task", ""))}</td>')
        lines.append(f'                                <td>{escape(action.get("owner", ""))}</td>')
        lines.append(f'                                <td>{escape(action.get("deadline", ""))}</td>')
        lines.append("                            </tr>")
    return "\n".join(lines)


def generate(data):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "..", "assets", "template.html")
    with open(template_path, "r", encoding="utf-8-sig") as f:
        template = f.read()

    core = data.get("core_conclusions", {})
    actions = data.get("actions", {})

    # Build footer line: meeting time + participants (if any) + generator tag
    footer_parts = []
    meeting_date = data.get("meeting_date", "").strip()
    if meeting_date:
        footer_parts.append(f"会议时间：{meeting_date}")
    participants = data.get("participants", "").strip()
    if participants:
        footer_parts.append(f"参会人员：{participants}")
    footer_parts.append("记录生成：AI自动整理")
    footer_line = " | ".join(footer_parts)

    replacements = {
        "{{TITLE}}": escape(data.get("title", "会议纪要")),
        "{{CONCLUSION_EMOJI}}": escape(core.get("emoji", "💡")),
        "{{CONCLUSION_TITLE}}": escape(core.get("title", "核心结论")),
        "{{CONCLUSION_ITEMS}}": build_conclusion_items(core.get("items", [])),
        "{{SECTIONS_TITLE}}": escape(data.get("sections_title", "会议讨论")),
        "{{CARDS}}": build_cards(data.get("cards", [])),
        "{{ACTION_EMOJI}}": escape(actions.get("emoji", "📋")),
        "{{ACTION_TITLE}}": escape(actions.get("title", "具体行动项与责任人")),
        "{{ACTION_SUBTITLE}}": escape(actions.get("subtitle", "")),
        "{{ACTION_ROWS}}": build_action_rows(actions.get("items", [])),
        "{{FOOTER_LINE}}": escape(footer_line),
    }

    for key, value in replacements.items():
        template = template.replace(key, value)

    output_path = data.get("output_path")
    if not output_path:
        output_path = os.path.join(os.getcwd(), "会议纪要.html")

    with open(output_path, "w", encoding="utf-8-sig") as f:
        f.write(template)

    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate meeting minutes HTML from JSON data.")
    parser.add_argument("json_file", help="Path to JSON input file")
    args = parser.parse_args()

    with open(args.json_file, "r", encoding="utf-8-sig") as f:
        data = json.load(f)

    output = generate(data)
    # Use sys.stdout with explicit encoding to avoid garbled Chinese characters on Windows
    sys.stdout.buffer.write(f"Generated: {output}\n".encode("utf-8"))


if __name__ == "__main__":
    main()
