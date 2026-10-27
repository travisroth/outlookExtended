# -*- coding: UTF-8 -*-
# NVDA add-on: Outlook Extended
# Copyright (C) 2026 Travis Roth
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

"""Finding where each message of an email chain begins, from the text of its paragraphs.

A reply or forward quotes the messages before it, each under a header: Outlook's "From:", "Sent:",
"To:", "Subject:" block, a line such as "-----Original Message-----", or a line such as "On Monday,
Jane wrote:". This finds those headers in a list of paragraph texts and says where the message
after each header begins, so the reader can move from message to message instead of through every
signature and header.

Nothing here uses NVDA, so it is unit tested on its own. `messageChainNavigation` reads the
paragraphs from the reading pane and moves the browse mode cursor.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from collections.abc import Sequence

FROM = "from"
SENT = "sent"
TO = "to"
CC = "cc"
SUBJECT = "subject"

HEADER_LABELS = {
	# English
	"from": FROM,
	"sent": SENT,
	"date": SENT,
	"to": TO,
	"cc": CC,
	"bcc": CC,
	"subject": SUBJECT,
	# French
	"de": FROM,
	"envoyé": SENT,
	"à": TO,
	"cci": CC,
	"objet": SUBJECT,
	# German
	"von": FROM,
	"gesendet": SENT,
	"datum": SENT,
	"an": TO,
	"betreff": SUBJECT,
	# Spanish
	"enviado": SENT,
	"fecha": SENT,
	"para": TO,
	"cco": CC,
	"asunto": SUBJECT,
}
"""Header labels as Outlook and other mail programs write them, lower case, to what each means.

The labels are in the language of whoever replied, not the reader's, so every language's are
looked for at once. "de" is From in both French and Spanish."""

HEADER_LOOKAHEAD = 8
"""How many paragraphs after a "From:" line the rest of its header block is looked for in."""

_LABEL_LINE = re.compile(r"^\s*([^\W\d_][^\W\d_ ]*(?: [^\W\d_]+)?)\s*:\s*(.*)$")
"""A "Label: value" line. The space before the colon is French typography."""

_SEPARATOR = re.compile(
	r"^\s*-{2,}\s*(original message|message d'origine|ursprüngliche nachricht|mensaje original"
	r"|forwarded message|message transféré|weitergeleitete nachricht|mensaje reenviado)\s*-{2,}\s*$"
	r"|^\s*begin forwarded message\s*:?\s*$",
	re.IGNORECASE,
)
"""A line that starts a quoted or forwarded message by itself."""

_WROTE = re.compile(
	r"^\s*(on\s.{3,200}\swrote|le\s.{3,200}\sa écrit|am\s.{3,200}\sschrieb.{0,100}|el\s.{3,200}\sescribió)\s*:\s*$",
	re.IGNORECASE,
)
"""The line a reply's quotation starts with in Gmail, Apple Mail and others: "On ... wrote:"."""

_WROTE_PARTS = re.compile(
	r"^\s*on\s+(?P<sent>.+?\d{1,2}:\d{2}(?::\d{2})?(?:\s*[ap]\.?m\.?)?)[,\s]+(?P<sender>.+?)\s+wrote\s*:\s*$",
	re.IGNORECASE,
)
"""An English "On <date and time>, <sender> wrote:" line, in Apple Mail's form and Gmail's."""

_RULE = re.compile(r"^\s*_{10,}\s*$")
"""The line of underscores Outlook puts above a reply's header."""

_LINE_BREAK = re.compile("[\r\n\x0b\x0c\u2028\u2029]")
"""What ends a line inside a paragraph: Word's manual line break is the vertical tab."""


@dataclass(frozen=True)
class Message:
	"""Where one message of a chain is, by paragraph index."""

	start: int
	"""The first paragraph of its header, or of the email for the message on top."""

	content: int
	"""The first paragraph after its header, where reading it starts."""

	sender: str | None = None
	"""Who it is from, as its header says, or None where the header does not say."""

	sent: str | None = None
	"""When it was sent, as its header says, or None."""

	header: str | None = None
	"""Its header line as written, where who and when could not be told apart in it, or None."""


def labelOf(text: str) -> tuple[str | None, str]:
	""":return: what a "Label: value" line's label means, or None, and its value."""
	match = _LABEL_LINE.match(text)
	if not match:
		return None, ""
	return HEADER_LABELS.get(match.group(1).strip().lower()), match.group(2).strip()


def _isBlank(text: str) -> bool:
	return not text.strip()


def linesOf(text: str) -> list[str]:
	""":return: the lines of a paragraph that are not blank.

	Outlook writes a reply's header as one paragraph with a line break after each field, so a
	paragraph is read a line at a time.
	"""
	return [line for line in _LINE_BREAK.split(text) if line.strip()]


def _wroteLine(paragraphs: Sequence[str], index: int) -> tuple[int, str]:
	""":return: how many paragraphs an "On ... wrote:" line at an index takes, 0 if there is none,
	and the line.

	It may wrap over two lines of one paragraph, or over two paragraphs, where the mail program
	broke the line.
	"""
	lines = linesOf(paragraphs[index])
	if not lines:
		return 0, ""
	for candidate in (lines[0].strip(), " ".join(line.strip() for line in lines)):
		if _WROTE.match(candidate):
			return 1, candidate
	if index + 1 < len(paragraphs):
		following = linesOf(paragraphs[index + 1])
		if following:
			candidate = f"{lines[-1].strip()} {following[0].strip()}"
			if _WROTE.match(candidate):
				return 2, candidate
	return 0, ""


def wroteLineParts(line: str) -> tuple[str | None, str | None]:
	""":return: who wrote and when, from an English "On ... wrote:" line, or (None, None).

	The date comes first and ends at its time, which is what tells it from the name after it.
	"""
	match = _WROTE_PARTS.match(line)
	if not match:
		return None, None
	# A line broken inside an address is joined with a space after its "<".
	sender = re.sub(r"<\s+", "<", match.group("sender").strip())
	return sender, match.group("sent").strip()


def _headerBlock(paragraphs: Sequence[str], index: int, skipLines: int = 0) -> tuple[int, dict[str, str]]:
	""":return: the index after a run of header lines and blanks from an index, and what they said.

	:param skipLines: how many lines of the first paragraph are not part of the header, such as a
		separator line written in the same paragraph as the fields under it.
	"""
	fields: dict[str, str] = {}
	end = index
	while end < len(paragraphs) and end < index + HEADER_LOOKAHEAD:
		lines = linesOf(paragraphs[end])
		if end == index:
			lines = lines[skipLines:]
		for line in lines:
			kind, value = labelOf(line)
			if kind is None or kind in fields:
				# Not a header line, or a label this header already has, which starts another.
				return end, fields
			fields[kind] = value
		end += 1
	return end, fields


def _isFullHeader(fields: dict[str, str]) -> bool:
	""":return: whether header fields are a message's: who from, when, and to whom or about what."""
	return FROM in fields and SENT in fields and (TO in fields or SUBJECT in fields)


def _skipBlanks(paragraphs: Sequence[str], index: int) -> int:
	while index < len(paragraphs) and _isBlank(paragraphs[index]):
		index += 1
	return index


def messageAt(paragraphs: Sequence[str], index: int) -> Message | None:
	""":return: the message whose header starts at a paragraph, or None if none does.

	A header is a separator line, an "On ... wrote:" line, a "From:" line followed within
	L{HEADER_LOOKAHEAD} paragraphs by a sent or date line and a to or subject line, or a line of
	underscores with such a header under it. A lone "From:" in someone's text or signature is not one.
	"""
	lines = linesOf(paragraphs[index])
	if not lines:
		return None
	first = lines[0]
	if _SEPARATOR.match(first):
		end, fields = _headerBlock(paragraphs, index, skipLines=1)
		return Message(index, _skipBlanks(paragraphs, end), fields.get(FROM), fields.get(SENT))
	if _RULE.match(first):
		end, fields = _headerBlock(paragraphs, index, skipLines=1)
		if not _isFullHeader(fields):
			return None
		return Message(index, _skipBlanks(paragraphs, end), fields.get(FROM), fields.get(SENT))
	wrote, line = _wroteLine(paragraphs, index)
	if wrote:
		sender, sent = wroteLineParts(line)
		header = None if sender else line
		return Message(index, _skipBlanks(paragraphs, index + wrote), sender, sent, header)
	kind, _value = labelOf(first)
	if kind != FROM:
		return None
	end, fields = _headerBlock(paragraphs, index)
	if not _isFullHeader(fields):
		return None
	return Message(index, _skipBlanks(paragraphs, end), fields.get(FROM), fields.get(SENT))


def nextMessage(paragraphs: Sequence[str], current: int) -> Message | None:
	""":return: the first message whose header starts after the paragraph the reader is in, or None.

	:param paragraphs: the texts of the paragraphs from the one the reader is in onwards, or more.
	:param current: the index of the paragraph the reader is in.
	"""
	for index in range(current + 1, len(paragraphs)):
		found = messageAt(paragraphs, index)
		if found is not None:
			return found
	return None


def previousMessage(paragraphs: Sequence[str], current: int, fromTheTop: bool) -> Message | None:
	""":return: the message before the one the reader is in, or None if there is none to go to.

	The message the reader is in is the last one whose header starts at or before their paragraph,
	so a reader in a header is in that header's message. Before every quoted message is the email
	itself, which starts at its first paragraph that is not blank.

	:param paragraphs: the texts of the paragraphs up to and including the one the reader is in.
	:param current: the index of the paragraph the reader is in.
	:param fromTheTop: whether the paragraphs start at the top of the email. If not, the email
		itself is not among them and is not offered.
	"""
	starts = []
	index = 0
	while index <= current and index < len(paragraphs):
		found = messageAt(paragraphs, index)
		if found is not None:
			starts.append(found)
			index = max(found.content, index + 1)
			continue
		index += 1
	inside = [message for message in starts if message.start <= current]
	if not inside:
		return None
	if len(inside) >= 2:
		return inside[-2]
	if not fromTheTop:
		return None
	return Message(0, _skipBlanks(paragraphs, 0))
