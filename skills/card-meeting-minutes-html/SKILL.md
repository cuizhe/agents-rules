---
name: card-meeting-minutes-html
description: "Generate structured meeting minutes as HTML with a fixed card+table layout. Structure: title → core conclusions → three-column cards (status/risk, business strategy, implementation path) → action items table. Use when the user asks to: create meeting minutes (会议纪要), generate meeting notes (会议记录), produce a meeting summary (会议总结), convert meeting transcript to HTML, save meeting record as HTML, or mentions keywords like '会议纪要', '会议记录', '会议总结', 'meeting minutes', 'meeting notes', 'meeting summary', 'meeting transcript'. Supports text transcripts, voice-to-text output, Doubao AI meeting content, and other meeting materials. Output is always saved as a .html file."
---

# 卡片式会议纪要

Generate structured meeting minutes as HTML with a fixed layout: title → core conclusions → three-column cards → action items table.

## Workflow

1. **Receive meeting content** from the user (text transcript, voice-to-text, Doubao AI meeting content, or manual notes).
2. **Extract meeting metadata** (title, date, participants):
   - **Participants — two extraction paths:**
     - **Path A (preferred):** Scan the transcript for speaker identifiers. If speakers are labeled with real names (e.g., `说话人A`、`说话人B`、`老丁`、`伟哥`), parse all unique names from the transcript and use them as the participant list.
     - **Path B (fallback):** If the transcript only uses generic placeholders (e.g., `说话人1`、`说话人2`、`Speaker 1`、`Speaker 2`), look for a separate participant list provided by the user — typically an image file (e.g., `参会人员名单_*.jpg`/`*.png`) or a text file. Read it and extract all names. If the user does not provide a list, ask for it before proceeding.
   - **Map nicknames to formal names:** Common nicknames in the transcript (e.g., "老丁", "建峰", "伟哥") must be mapped to their full names from the participant list. This mapping is used both in the `participants` field and throughout the content.
3. **Analyze and extract** structured information:
   - Core conclusions (3-5 key takeaways)
   - Three themed cards with bullet points:
     - Card 1: current status and risks (blue theme, ⚠️)
     - Card 2: business strategy (cyan theme, 🎯)
     - Card 3: implementation path (lime theme, </>)
   - Action items table (no, task, owner, deadline) — **see Action Items Extraction Checklist below**
4. **Normalize names across all content**: Replace every occurrence of nicknames (in task descriptions, owner fields, and card content) with the corresponding full names from the participant list. Use the person's actual title if the user has specified it (e.g., "交付主管" rather than "事业部总经理").
5. **Produce JSON data file** matching the schema below.
6. **Run the generation script**: `python scripts/generate_minutes.py <json_file>`
7. **Confirm output path** with the user.

## JSON Data Schema

```json
{
  "title": "Meeting title",
  "meeting_date": "2026年6月17日 14:00-15:30",
  "participants": "张三、李四、王五",
  "core_conclusions": {
    "title": "核心结论",
    "emoji": "💡",
    "items": ["conclusion 1", "conclusion 2"]
  },
  "sections_title": "Overall section title (e.g. 'AI编码应用策略与风险')",
  "cards": [
    {
      "title": "Card title",
      "emoji": "⚠️",
      "color": "blue",
      "subtitle": "One-line summary",
      "items": ["detail 1", "detail 2"]
    }
  ],
  "actions": {
    "title": "具体行动项与责任人",
    "emoji": "📋",
    "subtitle": "One-line summary",
    "items": [
      {"no": 1, "task": "...", "owner": "...", "deadline": "..."}
    ]
  },
  "output_path": "C:/path/to/output.html"
}
```

**Card color values**: `blue` (status/risk), `cyan` (strategy), `lime` (implementation).

## Action Items Extraction Checklist

Action items come from **two independent dimensions**. Do not merge them — each dimension can produce its own tasks.

| Dimension | Source | Example trigger phrases |
|---|---|---|
| **Expert/Leader Suggestions** | 专家建议、领导指示、策略要求 | "建议..."、"可以尝试..."、"大胆的..."、"要..." |
| **Meeting Natural Legacy Items** | 会议自然产生的遗留问题、对话中自然安排的待办 | "做个遗留问题"、"下次约个时间"、"XX记一下"、"XX跟XX一起看"、"到时候再开一次专题会议" |

**Extraction rules:**
1. **Segmented scan**: Scan the **last 20% of the transcript** (meeting closing section) specifically for phrases like: `遗留`, `下次`, `约时间`, `记一下`, `跟XX一起`, `专题会议`, `待办`.
2. **Distinguish from existing topics**: A legacy item about the same topic as an existing action item is **NOT a duplicate** if the trigger, owner, or goal is different. For example:
   - Action item: "伟哥建议用AI辅助迁移" (expert suggestion) → proactive
   - Legacy item: "李剑锋让崔哲记录遗留问题，下次专题会议" (meeting legacy) → reactive follow-up
3. **Owner extraction**: When a speaker says `"XX记一下"` or `"XX跟XX一起"` or `"XX到时候记一下"`, the person named (XX) is the owner. The recorder/tracker is a distinct role — e.g., `"崔哲记一下"` → owner is "崔哲（记录）".
4. **Mark as [遗留事项]**: Prefix the task description with `[遗留事项]` so it is visually distinct from proactive action items.

**Name normalization rules**:
- If the transcript contains real speaker names (Path A): use those names as-is for the participant list, but still map any nicknames to full names in the content.
- If the transcript only contains generic placeholders (Path B): read the user-provided participant list (image or text file) to extract all full names, then build a nickname→full-name mapping.
- Replace all occurrences in task descriptions, owner fields, and card content. Use the person's actual title if the user has specified it (e.g., "交付主管" rather than "事业部总经理").

## Resources

- `scripts/generate_minutes.py` — reads JSON and produces the final HTML file using `assets/template.html`. The script uses `sys.stdout.buffer` with UTF-8 encoding to prevent garbled console output on Windows.
- `assets/template.html` — self-contained HTML template with all CSS inlined. No external CDN dependencies (no Tailwind CSS). Uses a soft, professional color scheme matching the reference design: light blue, light cyan, light green, and light orange. The action table has an 80px-wide "序号" column to prevent text wrapping.
- `references/example-minutes-screenshot.png` — reference screenshot showing the intended visual style (card layout, soft colors, action table). Load and review this image to understand the expected output aesthetics before generating.

## Design Notes

- The template is **fully self-contained** — all CSS is inlined, no external network requests. This ensures correct rendering even in offline environments or corporate networks with CDN restrictions.
- **Color scheme**: Core conclusions use light blue (#eef6ff), risk cards use soft cyan (#e6f7ff), strategy cards use soft teal (#e6fffb), implementation cards use soft green (#f6ffed), and action items use soft orange (#fff7e6).
- **Footer**: Automatically combines `meeting_date` and `participants` (if provided) into a single footer line, separated by " | ".
- **Reference example**: The file `references/example-minutes-screenshot.png` shows the intended visual output style. When generating the first time for a new project, load and review it to ensure the generated HTML matches the card layout, color palette, and table spacing shown in the reference.
- **HTML structure alignment**: `generate_minutes.py` generates cards with `.card-header` and `.card-inner` classes. The template CSS must match this structure. If modifying either the script or the template, ensure both sides remain aligned — mismatched selectors (e.g., using `.card > h3` when the template expects `.card-header h3`) will cause style failures (e.g., wrong font size, missing padding, or default browser styles overriding intended ones).
- **Card typography tuning**: Card list items use `font-size: 13.5px` and `line-height: 1.8` for comfortable reading. The card grid uses `gap: 28px` and card padding is `24px` to prevent crowding on dense content. Adjust these values together with the `example-minutes-screenshot.png` as the visual reference.
