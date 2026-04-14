from typing import Optional

import gradio as gr
from ktem.app import BasePage
from ktem.db.models import IssueReport, engine
from sqlmodel import Session


class ReportIssue(BasePage):
    def __init__(self, app):
        self._app = app
        self.on_building_ui()

    def on_building_ui(self):
        with gr.Accordion(label="反馈", open=False, elem_id="report-accordion"):
            self.correctness = gr.Radio(
                choices=[
                    ("回答正确", "correct"),
                    ("回答错误", "incorrect"),
                ],
                label="准确性：",
            )
            self.issues = gr.CheckboxGroup(
                choices=[
                    ("内容不当", "offensive"),
                    ("引用有误", "wrong-evidence"),
                ],
                label="其他问题：",
            )
            self.more_detail = gr.Textbox(
                placeholder=(
                    "请描述具体问题（例如：哪里出错了，正确答案是什么， "
                    "等）"
                ),
                container=False,
                lines=3,
            )
            gr.Markdown(
                "将发送当前对话和用户设置以协助排查"
                "help with investigation"
            )
            self.report_btn = gr.Button("提交反馈")

    def report(
        self,
        correctness: str,
        issues: list[str],
        more_detail: str,
        conv_id: str,
        chat_history: list,
        settings: dict,
        user_id: Optional[int],
        info_panel: str,
        chat_state: dict,
        *selecteds,
    ):
        selecteds_ = {}
        for index in self._app.index_manager.indices:
            if index.selector is not None:
                if isinstance(index.selector, int):
                    selecteds_[str(index.id)] = selecteds[index.selector]
                elif isinstance(index.selector, tuple):
                    selecteds_[str(index.id)] = [selecteds[_] for _ in index.selector]
                else:
                    print(f"Unknown selector type: {index.selector}")

        with Session(engine) as session:
            issue = IssueReport(
                issues={
                    "correctness": correctness,
                    "issues": issues,
                    "more_detail": more_detail,
                },
                chat={
                    "conv_id": conv_id,
                    "chat_history": chat_history,
                    "info_panel": info_panel,
                    "chat_state": chat_state,
                    "selecteds": selecteds_,
                },
                settings=settings,
                user=user_id,
            )
            session.add(issue)
            session.commit()
        gr.Info("Thank you for your feedback")
