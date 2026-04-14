import os
import tempfile
from datetime import datetime
from pathlib import Path

import gradio as gr
from ktem.app import BasePage


def generate_report_doc(
    exp_name,
    exp_date,
    exp_person,
    exp_purpose,
    data_file,
    image_files,
):
    """生成实验报告 Word 文档"""
    try:
        from docx import Document
        from docx.shared import Inches, Pt
        from openai import OpenAI
    except ImportError as e:
        return None, f"缺少依赖库：{e}，请执行 pip install python-docx openai"

    if not exp_name:
        return None, "请填写实验名称"

    # ── 1. 解析上传的数据文件 ──────────────────────────
    data_text = ""
    if data_file is not None:
        try:
            file_path = data_file.name if hasattr(data_file, "name") else data_file
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                data_text = f.read()
        except Exception as e:
            data_text = f"（数据文件读取失败：{e}）"

    # ── 2. 调用大模型生成分析摘要 ──────────────────────
    analysis_text = ""
    llm_base_url = os.getenv("REPORT_LLM_BASE_URL", "")
    llm_api_key = os.getenv("REPORT_LLM_API_KEY", "")
    llm_model = os.getenv("REPORT_LLM_MODEL", "")

    if llm_base_url and llm_api_key and llm_model and data_text:
        try:
            client = OpenAI(base_url=llm_base_url, api_key=llm_api_key)
            prompt = f"""你是一名专业的实验分析工程师。
以下是一份仿真实验的数据结果，请根据数据内容生成一段专业的实验结果分析与总结。
要求：
1. 概括实验的主要结果和关键指标
2. 分析数据反映的规律或问题
3. 给出简要的结论和建议
4. 语言专业、简洁，200-400字

实验名称：{exp_name}
实验目的：{exp_purpose}

实验数据：
{data_text[:3000]}
"""
            resp = client.chat.completions.create(
                model=llm_model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=600,
            )
            analysis_text = resp.choices[0].message.content.strip()
        except Exception as e:
            analysis_text = f"（大模型分析生成失败：{e}，以下为数据摘要）\n{data_text[:500]}"
    else:
        if data_text:
            analysis_text = "（未配置大模型接口，请在 .env 中设置 REPORT_LLM_BASE_URL / REPORT_LLM_API_KEY / REPORT_LLM_MODEL）"
        else:
            analysis_text = "（未上传数据文件，无法生成分析）"

    # ── 3. 用 python-docx 生成 Word 文档 ──────────────
    doc = Document()

    # 标题
    title = doc.add_heading(f"{exp_name} 实验仿真报告", level=0)
    title.alignment = 1  # 居中

    doc.add_paragraph("")

    # 第一章：基本信息
    doc.add_heading("第一章  实验基本信息", level=1)
    info_table = doc.add_table(rows=4, cols=2)
    info_table.style = "Table Grid"
    labels = ["实验名称", "实验日期", "实验人员", "实验目的"]
    values = [
        exp_name,
        exp_date or datetime.now().strftime("%Y-%m-%d"),
        exp_person or "未填写",
        exp_purpose or "未填写",
    ]
    for i, (label, value) in enumerate(zip(labels, values)):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[1].text = value

    doc.add_paragraph("")

    # 第二章：实验数据
    doc.add_heading("第二章  实验数据", level=1)
    if data_text:
        # 只展示前2000字符，避免文档过大
        display_data = data_text[:2000]
        if len(data_text) > 2000:
            display_data += "\n...\n（数据过长，已截断，完整数据见原始文件）"
        doc.add_paragraph(display_data).style.font.size = Pt(9)
    else:
        doc.add_paragraph("（未上传数据文件）")

    doc.add_paragraph("")

    # 第三章：图表
    doc.add_heading("第三章  实验图表", level=1)
    if image_files:
        files = image_files if isinstance(image_files, list) else [image_files]
        added = 0
        for img in files:
            try:
                img_path = img.name if hasattr(img, "name") else img
                doc.add_picture(img_path, width=Inches(5.5))
                doc.add_paragraph(f"图 {added+1}")
                added += 1
            except Exception as e:
                doc.add_paragraph(f"（图片插入失败：{e}）")
    else:
        doc.add_paragraph("（未上传图表）")

    doc.add_paragraph("")

    # 第四章：结果分析与总结
    doc.add_heading("第四章  结果分析与总结", level=1)
    doc.add_paragraph(analysis_text)

    doc.add_paragraph("")

    # 页脚信息
    doc.add_paragraph(
        f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ).italic = True

    # ── 4. 保存文件 ────────────────────────────────────
    safe_name = "".join(c for c in exp_name if c.isalnum() or c in "._- ").strip()
    filename = f"{safe_name}_实验报告_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    output_path = os.path.join(tempfile.gettempdir(), filename)
    doc.save(output_path)

    return output_path, "✅ 报告生成成功，请点击下方链接下载"


class ReportPage(BasePage):
    """报告生成页面"""

    def __init__(self, app):
        self._app = app
        self.on_building_ui()

    def on_building_ui(self):
        with gr.Row():
            with gr.Column(scale=1):
                gr.Markdown("## 📋 实验信息填写")
                self.exp_name = gr.Textbox(
                    label="实验名称 *", placeholder="请输入实验名称"
                )
                self.exp_date = gr.Textbox(
                    label="实验日期",
                    placeholder=datetime.now().strftime("%Y-%m-%d"),
                )
                self.exp_person = gr.Textbox(
                    label="实验人员", placeholder="请输入实验人员姓名"
                )
                self.exp_purpose = gr.Textbox(
                    label="实验目的",
                    placeholder="请简要描述实验目的",
                    lines=3,
                )
                gr.Markdown("## 📁 文件上传")
                self.data_file = gr.File(
                    label="上传仿真结果文件（CSV / JSON / TXT）",
                    file_types=[".csv", ".json", ".txt", ".log"],
                )
                self.image_files = gr.File(
                    label="上传实验图表（可多选，PNG / JPG）",
                    file_count="multiple",
                    file_types=[".png", ".jpg", ".jpeg"],
                )
                self.generate_btn = gr.Button(
                    "🚀 生成报告", variant="primary", size="lg"
                )

            with gr.Column(scale=1):
                gr.Markdown("## 📄 生成结果")
                self.status_box = gr.Textbox(
                    label="状态", interactive=False, lines=2
                )
                self.output_file = gr.File(
                    label="📥 下载报告", interactive=False
                )
                gr.Markdown(
                    """
### 使用说明
1. 填写实验基本信息（实验名称为必填项）
2. 上传仿真结果数据文件（支持 CSV/JSON/TXT）
3. 可选：上传实验截图或图表
4. 点击「生成报告」按钮
5. 等待生成完成后下载 Word 文档

### 报告包含内容
- 第一章：实验基本信息
- 第二章：实验数据
- 第三章：实验图表
- 第四章：AI 分析与总结
                    """
                )

    def on_register_events(self):
        self.generate_btn.click(
            fn=generate_report_doc,
            inputs=[
                self.exp_name,
                self.exp_date,
                self.exp_person,
                self.exp_purpose,
                self.data_file,
                self.image_files,
            ],
            outputs=[self.output_file, self.status_box],
        )
