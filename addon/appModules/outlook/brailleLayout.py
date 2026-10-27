# -*- coding: UTF-8 -*-
# NVDA add-on: Outlook Extended
# Copyright (C) 2026 Travis Roth
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

"""A braille layout for the BrlMultiline add-on: the message list beside the reading pane.

BrlMultiline asks the app module for a layout on each focus change, through a method it finds by
name, so nothing here imports from that add-on and Outlook works the same without it. While the
focus is in the message list, one segment follows the focus and the other shows the body of the
message in the reading pane. Moving into the message takes the layout down, so the whole display
reads it.

Outlook keeps one document window for every message it shows in the reading pane, so NVDA's
objects for two messages compare equal. The object given to BrlMultiline carries a key made from
the reading pane's header fields, which tells BrlMultiline that a new message is shown and should
be read from its top.
"""

# The annotations below are not evaluated, for the Python of older NVDA versions.
from __future__ import annotations

import ctypes

import winUser
from logHandler import log
from NVDAObjects.window import Window
from windowUtils import WNDENUMPROC

LAYOUT_NAME = "mail"

MIN_CELLS = 20
"""The narrowest one row display that is split between the list and the message."""

MESSAGE_LIST_CLASS = "OutlookGrid"
MESSAGE_BODY_CLASS = "_WwG"
HEADER_FIELD_CLASS = "RichEdit20WPT"
HEADER_DIALOG_CLASS = "#32770"

SHOW_KEY_ATTRIBUTE = "brlMultilineShowKey"
"""Read by BrlMultiline to tell apart two messages that NVDA sees as the same object."""

HEADER_TEXT_LENGTH = 256
SEND_MESSAGE_TIMEOUT_MS = 100
WM_GETTEXT = 0x000D


def isInMessageList(obj) -> bool:
	"""Is this object in the message list, a message or a group heading?"""
	try:
		return obj.windowClassName == MESSAGE_LIST_CLASS
	except Exception:
		return False


def _visibleDescendants(parent: int, className: str) -> list[int]:
	"""The visible windows of a class below a window, in z order."""
	found = []

	@WNDENUMPROC
	def callback(window, data):
		if winUser.getClassName(window) == className and winUser.isWindowVisible(window):
			found.append(window)
		return True

	ctypes.windll.user32.EnumChildWindows(parent, callback, 0)
	return found


def _windowText(window: int) -> str:
	"""The text of a window in Outlook, without waiting on Outlook if it is busy."""
	buffer = ctypes.create_unicode_buffer(HEADER_TEXT_LENGTH)
	result = ctypes.c_size_t()
	if not ctypes.windll.user32.SendMessageTimeoutW(
		window,
		WM_GETTEXT,
		HEADER_TEXT_LENGTH,
		buffer,
		winUser.SMTO_ABORTIFHUNG,
		SEND_MESSAGE_TIMEOUT_MS,
		ctypes.byref(result),
	):
		return ""
	return buffer.value


def _headerDialog(bodyWindow: int) -> int | None:
	"""The dialog holding the reading pane's header fields and its body."""
	window = bodyWindow
	while window:
		window = winUser.getAncestor(window, winUser.GA_PARENT)
		if window and winUser.getClassName(window) == HEADER_DIALOG_CLASS:
			return window
	return None


def messageKey(bodyWindow: int) -> tuple[str, ...] | None:
	"""Name the message the reading pane shows, by the text of its header fields.

	From, To, Subject, Sent and the rest, whatever control IDs this version of Outlook gives them.
	None when there are none to read, and BrlMultiline then compares the objects as it would anyway.
	"""
	dialog = _headerDialog(bodyWindow)
	if dialog is None:
		return None
	key = tuple(_windowText(field) for field in _visibleDescendants(dialog, HEADER_FIELD_CLASS))
	return key if any(key) else None


def findMessageBody(focus):
	"""The body of the message in the reading pane of the focus's window, or None."""
	root = winUser.getAncestor(focus.windowHandle, winUser.GA_ROOT)
	if not root:
		return None
	bodies = _visibleDescendants(root, MESSAGE_BODY_CLASS)
	if not bodies:
		return None
	body = Window(windowHandle=bodies[0])
	key = messageKey(bodies[0])
	if key is not None:
		setattr(body, SHOW_KEY_ATTRIBUTE, key)
	return body


class MessageListLayout:
	"""Offers BrlMultiline the message list beside the reading pane. One per app module."""

	def __init__(self, getFocus):
		"""
		:param getFocus: returns the object with the focus now, asked each time the reading pane is.
		"""
		super().__init__()
		self._getFocus = getFocus

	def layout(self, focus, area):
		"""Answer `brlMultilineLayout`: the layout while the focus is in the message list, else None.

		:param focus: the object with the focus.
		:param area: the free part of the display, with ``rows`` and ``columns``.
		"""
		if not isInMessageList(focus):
			return None
		if area.rows > 1:
			# A message row holds sender, subject, date and more, so two rows where there are
			# rows to spare.
			focusSize = 2 if area.rows >= 6 else 1
		else:
			if area.columns < MIN_CELLS:
				return None
			focusSize = area.columns // 2
		return {
			"name": LAYOUT_NAME,
			"segments": [
				# The flow, where the reader has it, shows the messages around the focus rather
				# than one row of the focused message. BrlMultiline ignores it on one row.
				{"size": focusSize, "show": "focus", "flow": True},
				{"show": self.messageBody},
			],
		}

	def messageBody(self):
		"""The reading pane's message, asked again on each focus change and about once a second."""
		focus = self._getFocus()
		if focus is None or not isInMessageList(focus):
			return None
		try:
			return findMessageBody(focus)
		except Exception:
			log.debugWarning("Could not find the reading pane for the braille layout", exc_info=True)
			return None
