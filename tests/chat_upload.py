"""Test helper: attach files through the chat bar's + button, which AppTest cannot click."""

from hashlib import sha256
from unittest.mock import call, patch

import streamlit as st
from streamlit.elements.widgets.chat import ChatInputValue

REAL_CHAT_INPUT = st.chat_input


class ChatUpload:
    """Wraps st.chat_input and attaches the provided upload whenever it changes.

    ``return_value``/``side_effect`` work like the old ``st.file_uploader`` patches: the side effect is
    called with ``key="excel_upload_0"`` on every run, and a new or changed file is attached on that run.
    """

    def __init__(self, return_value=None, side_effect=None):
        self.return_value = return_value
        self.side_effect = side_effect
        self.call_args_list = []

    def current(self):
        if self.side_effect is not None:
            return self.side_effect(key="excel_upload_0")
        return self.return_value

    def __call__(self, *args, **kwargs):
        self.call_args_list.append(call(*args, **kwargs))
        value = REAL_CHAT_INPUT(*args, **kwargs)
        upload = self.current()
        if upload is None:
            return value
        try:
            marker = (id(upload), sha256(upload.getvalue()).hexdigest())
        except OSError:
            marker = (id(upload), None)
        if marker == st.session_state.get("_test_attached_upload"):
            return value
        st.session_state["_test_attached_upload"] = marker
        text = value if isinstance(value, str) else getattr(value, "text", "") or ""
        return ChatInputValue(text=text, files=[upload], _include_files=True)


def patch_chat_upload(return_value=None, side_effect=None):
    return patch("streamlit.chat_input", new=ChatUpload(return_value, side_effect))
