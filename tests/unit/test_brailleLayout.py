# -*- coding: UTF-8 -*-
# tests/unit/test_brailleLayout.py
# NVDA add-on: Outlook Extended
# Copyright (C) 2026 Travis Roth
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

"""Tests for the BrlMultiline layout: the message list beside the reading pane.

Outlook's windows are a fake tree, so these run without NVDA or Outlook. Run with the other unit
tests from tests/unit/__init__.py, which puts the add-on's source on the path.
"""

import sys
import types
import unittest

# The window tree each test builds: handle to (class, parent, text, visible).
WINDOWS = {}


def _stubModules():
	winUser = types.ModuleType("winUser")
	winUser.GA_PARENT = 1
	winUser.GA_ROOT = 2
	winUser.SMTO_ABORTIFHUNG = 2

	def getAncestor(window, flags):
		parent = WINDOWS[window][1] if window in WINDOWS else 0
		if flags == winUser.GA_ROOT:
			while parent and WINDOWS[parent][1]:
				parent = WINDOWS[parent][1]
		return parent

	winUser.getAncestor = getAncestor
	winUser.getClassName = lambda window: WINDOWS[window][0]
	winUser.isWindowVisible = lambda window: WINDOWS[window][3]
	sys.modules.setdefault("winUser", winUser)

	windowUtils = types.ModuleType("windowUtils")
	windowUtils.WNDENUMPROC = lambda function: function
	sys.modules.setdefault("windowUtils", windowUtils)

	class Window:
		def __init__(self, windowHandle):
			self.windowHandle = windowHandle

	nvdaObjects = types.ModuleType("NVDAObjects")
	nvdaObjectsWindow = types.ModuleType("NVDAObjects.window")
	nvdaObjectsWindow.Window = Window
	nvdaObjects.window = nvdaObjectsWindow
	sys.modules.setdefault("NVDAObjects", nvdaObjects)
	sys.modules.setdefault("NVDAObjects.window", nvdaObjectsWindow)


_stubModules()

import brailleLayout  # noqa: E402 - Needed here after the stubs.


def _descendants(parent, className):
	"""What EnumChildWindows would find, from the fake tree."""

	def below(window):
		ancestor = WINDOWS[window][1]
		while ancestor:
			if ancestor == parent:
				return True
			ancestor = WINDOWS[ancestor][1]
		return False

	return [
		window
		for window, (cls, _parent, _text, visible) in WINDOWS.items()
		if cls == className and visible and below(window)
	]


brailleLayout._visibleDescendants = _descendants
brailleLayout._windowText = lambda window: WINDOWS[window][2]


def _mainWindow(subject="First message", withReadingPane=True):
	"""Outlook's main window as it was found in Outlook 365 classic: list, then reading pane."""
	WINDOWS.clear()
	WINDOWS[1] = ("rctrl_renwnd32", 0, "Inbox - Outlook", True)
	WINDOWS[10] = ("OutlookGrid", 1, "Table View", True)
	if withReadingPane:
		WINDOWS[20] = ("rctrl_renwnd32", 1, "", True)
		WINDOWS[21] = ("#32770", 20, "", True)
		WINDOWS[22] = ("RichEdit20WPT", 21, subject, True)
		WINDOWS[23] = ("RichEdit20WPT", 21, "Sender", True)
		WINDOWS[24] = ("AfxWndA", 21, "", True)
		WINDOWS[25] = ("_WwB", 24, "Document1", True)
		WINDOWS[26] = ("_WwG", 25, "Message", True)


class _Focus:
	def __init__(self, windowHandle, windowClassName):
		self.windowHandle = windowHandle
		self.windowClassName = windowClassName


MONARCH = types.SimpleNamespace(rows=10, columns=32)
FOCUS80 = types.SimpleNamespace(rows=1, columns=80)


class TestMessageListLayout(unittest.TestCase):
	def setUp(self):
		_mainWindow()
		self.focus = _Focus(10, "OutlookGrid")
		self.layout = brailleLayout.MessageListLayout(lambda: self.focus)

	def test_offeredFromTheMessageList(self):
		layout = self.layout.layout(self.focus, MONARCH)
		self.assertEqual(layout["name"], "mail")
		self.assertEqual(layout["segments"][0], {"size": 2, "show": "focus", "flow": True})
		self.assertEqual(layout["segments"][1], {"show": self.layout.messageBody})

	def test_notOfferedInTheMessage(self):
		self.assertIsNone(self.layout.layout(_Focus(26, "_WwG"), MONARCH))

	def test_oneRowForTheFocusOnASmallArea(self):
		layout = self.layout.layout(self.focus, types.SimpleNamespace(rows=4, columns=32))
		self.assertEqual(layout["segments"][0]["size"], 1)

	def test_halfTheCellsOnOneRow(self):
		layout = self.layout.layout(self.focus, FOCUS80)
		self.assertEqual(layout["segments"][0]["size"], 40)

	def test_notOfferedOnTooFewCells(self):
		self.assertIsNone(self.layout.layout(self.focus, types.SimpleNamespace(rows=1, columns=14)))

	def test_theReadingPaneBodyIsShown(self):
		body = self.layout.messageBody()
		self.assertEqual(body.windowHandle, 26)

	def test_theBodyIsNamedByItsHeaders(self):
		self.assertEqual(self.layout.messageBody().brlMultilineShowKey, ("First message", "Sender"))

	def test_anotherMessageInTheSameWindowHasAnotherKey(self):
		first = self.layout.messageBody()
		_mainWindow(subject="Second message")
		second = self.layout.messageBody()
		self.assertEqual(first.windowHandle, second.windowHandle)
		self.assertNotEqual(first.brlMultilineShowKey, second.brlMultilineShowKey)

	def test_noReadingPaneBlanksTheSegment(self):
		_mainWindow(withReadingPane=False)
		self.assertIsNone(self.layout.messageBody())

	def test_aHiddenBodyBlanksTheSegment(self):
		"""Outlook hides the body for a moment while it changes messages."""
		WINDOWS[26] = ("_WwG", 25, "Message", False)
		self.assertIsNone(self.layout.messageBody())

	def test_noHeadersLeavesTheBodyUnkeyed(self):
		for window in (22, 23):
			WINDOWS[window] = ("RichEdit20WPT", 21, "", True)
		self.assertFalse(hasattr(self.layout.messageBody(), "brlMultilineShowKey"))

	def test_nothingOnceTheFocusLeavesTheList(self):
		self.focus = _Focus(26, "_WwG")
		self.assertIsNone(self.layout.messageBody())

	def test_aFailureBlanksTheSegment(self):
		self.focus = _Focus(999, "OutlookGrid")

		def broken(window, flags):
			raise OSError("Outlook went away")

		original = sys.modules["winUser"].getAncestor
		sys.modules["winUser"].getAncestor = broken
		try:
			self.assertIsNone(self.layout.messageBody())
		finally:
			sys.modules["winUser"].getAncestor = original


if __name__ == "__main__":
	unittest.main()
