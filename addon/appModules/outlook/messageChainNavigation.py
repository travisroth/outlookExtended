# -*- coding: UTF-8 -*-
# NVDA add-on: Outlook Extended
# Copyright (C) 2026 Travis Roth
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

"""Moving the browse mode cursor in a message from one message of its email chain to the next.

n and Shift+N, in the browse mode of a message being read. The keys belong to the browse mode
document rather than to the app module, so they apply only in browse mode and n types as usual
everywhere else. They are given to it by `ReadingPaneDocument`, an overlay on the message's Word
document object, which hands NVDA the usual browse mode class with L{MessageChainNavigation} mixed
in. Which paragraphs begin a message is decided by `messageChain`.
"""

import time

import addonHandler
import textInfos
import ui
from baseObject import ScriptableObject
from logHandler import log
from NVDAObjects import NVDAObject
from scriptHandler import script

from . import messageChain

try:
	from speech import speech
except ImportError:
	# Older versions of NVDA such as 2019.3.
	import speech


def _quickNavReason():
	""":return: why the cursor moves when quick navigation moves it, or None on NVDA too old to say."""
	try:
		from controlTypes import OutputReason
	except ImportError:
		return None
	return getattr(OutputReason, "QUICKNAV", None)


QUICKNAV_REASON = _quickNavReason()

addonHandler.initTranslation()

ADDON_SUMMARY = addonHandler.getCodeAddon().manifest["summary"]

SEARCH_SECONDS = 0.5
"""How long a search may read paragraphs before it gives up, since each is a call into Outlook."""

MAX_PARAGRAPHS = 3000
"""How many paragraphs a search reads at most, whatever the time."""


class MessageChainNavigation(ScriptableObject):
	"""Mixed into a message's browse mode document: n and Shift+N move between messages of its chain."""

	@script(
		# Translators: Input help message for a command in a message being read in Outlook.
		description=_("Moves to the next message in an email chain, past its header"),
		category=ADDON_SUMMARY,
		gesture="kb:n",
	)
	def script_nextChainMessage(self, gesture):
		self._moveInChain(forward=True)

	@script(
		# Translators: Input help message for a command in a message being read in Outlook.
		description=_("Moves to the previous message in an email chain, past its header"),
		category=ADDON_SUMMARY,
		gesture="kb:shift+n",
	)
	def script_previousChainMessage(self, gesture):
		self._moveInChain(forward=False)

	def _moveInChain(self, forward: bool) -> None:
		try:
			if forward:
				texts, positions = self._readParagraphs()
				found = messageChain.nextMessage(texts, 0)
			else:
				before, reachedTop = self._readParagraphsBack()
				# A little past the cursor too, so a header the cursor is in is read whole.
				after, afterPositions = self._readParagraphs(limit=messageChain.HEADER_LOOKAHEAD)
				texts = [text for text, _info in before] + after[1:]
				positions = [info for _text, info in before] + afterPositions[1:]
				found = messageChain.previousMessage(texts, len(before) - 1, fromTheTop=reachedTop)
		except Exception:
			log.debugWarning("Could not read the paragraphs of a message", exc_info=True)
			found = None
		if found is None or found.content >= len(positions):
			if forward:
				# Translators: Reported when there is no further message in an email chain.
				ui.message(_("No next message"))
			else:
				# Translators: Reported when there is no earlier message in an email chain.
				ui.message(_("No previous message"))
			return
		self._moveTo(positions[found.content], found)

	def _readParagraphs(self, limit: int = MAX_PARAGRAPHS):
		""":return: the texts of the paragraphs from the one the cursor is in onwards, and where each is."""
		info = self.selection.copy()
		info.expand(textInfos.UNIT_PARAGRAPH)
		texts = [info.text]
		positions = [info.copy()]
		deadline = time.monotonic() + SEARCH_SECONDS
		while len(texts) < limit and time.monotonic() < deadline:
			step = info.copy()
			step.collapse(end=True)
			step.expand(textInfos.UNIT_PARAGRAPH)
			if step.compareEndPoints(info, "startToStart") <= 0:
				# The end of the document, where the next paragraph is the last one again.
				break
			info = step
			texts.append(info.text)
			positions.append(info.copy())
		return texts, positions

	def _readParagraphsBack(self):
		""":return: the paragraphs up to the one the cursor is in, in document order, as (text, position),
		and whether they reach the top of the message."""
		info = self.selection.copy()
		info.expand(textInfos.UNIT_PARAGRAPH)
		found = [(info.text, info.copy())]
		deadline = time.monotonic() + SEARCH_SECONDS
		reachedTop = False
		while True:
			if len(found) >= MAX_PARAGRAPHS or time.monotonic() >= deadline:
				break
			step = info.copy()
			step.collapse()
			if step.move(textInfos.UNIT_PARAGRAPH, -1) == 0:
				reachedTop = True
				break
			step.expand(textInfos.UNIT_PARAGRAPH)
			if step.compareEndPoints(info, "startToStart") >= 0:
				reachedTop = True
				break
			info = step
			found.append((info.text, info.copy()))
		found.reverse()
		return found, reachedTop

	def _moveTo(self, paragraph, found) -> None:
		"""Put the browse mode cursor at the start of a paragraph, and say what message it is in."""
		target = paragraph.copy()
		target.collapse()
		if QUICKNAV_REASON is None:
			self.selection = target
		else:
			try:
				self._set_selection(target, reason=QUICKNAV_REASON)
			except TypeError:
				# A browse mode document whose selection takes no reason.
				self.selection = target
		if found.sender and found.sent:
			# Translators: Reported on moving to a message in an email chain: who it is from and when.
			ui.message(_("From {sender}, {sent}").format(sender=found.sender, sent=found.sent))
		elif found.sender:
			# Translators: Reported on moving to a message in an email chain: who it is from.
			ui.message(_("From {sender}").format(sender=found.sender))
		elif found.header:
			ui.message(found.header)
		elif found.start == 0:
			# Translators: Reported on moving back to the top of an email, before the messages it quotes.
			ui.message(_("Top of the message"))
		line = target.copy()
		line.expand(textInfos.UNIT_LINE)
		if QUICKNAV_REASON is not None:
			speech.speakTextInfo(line, unit=textInfos.UNIT_LINE, reason=QUICKNAV_REASON)
		else:
			speech.speakTextInfo(line, unit=textInfos.UNIT_LINE)


_withNavigation = {}
"""The browse mode classes with L{MessageChainNavigation} mixed in, by the class they extend."""


def withChainNavigation(base):
	""":return: a browse mode document class with message chain navigation, made once per class."""
	cls = _withNavigation.get(base)
	if cls is None:
		cls = _withNavigation[base] = type(
			f"{base.__name__}WithMessageChain", (MessageChainNavigation, base), {}
		)
	return cls


class ReadingPaneDocument(NVDAObject):
	"""A message's Word document, whose browse mode has message chain navigation."""

	def _get_treeInterceptorClass(self):
		return withChainNavigation(super().treeInterceptorClass)
