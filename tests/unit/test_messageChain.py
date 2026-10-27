# -*- coding: UTF-8 -*-
# tests/unit/test_messageChain.py
# NVDA add-on: Outlook Extended
# Copyright (C) 2026 Travis Roth
# This file is covered by the GNU General Public License.
# See the file COPYING.txt for more details.

"""Tests for finding the messages of an email chain in its paragraphs.

Run with the other unit tests from tests/unit/__init__.py, which puts the add-on's source on the
path.
"""

import unittest

from messageChain import Message, labelOf, messageAt, nextMessage, previousMessage

OUTLOOK_CHAIN = [
	"Sounds good, see you then.",  # 0
	"",
	"Travis",
	"",
	"From: Jane Doe <jane@example.com>",  # 4
	"Sent: Tuesday, October 21, 2026 3:15 PM",
	"To: Travis Roth <travis@example.com>",
	"Subject: RE: Lunch",
	"",
	"Tuesday works for me.",  # 9
	"Jane Doe",
	"Director of Things",
	"From: the desk of Jane",
	"",
	"From: Travis Roth <travis@example.com>",  # 14
	"Sent: Monday, October 20, 2026 9:02 AM",
	"To: Jane Doe <jane@example.com>",
	"Cc: Sam <sam@example.com>",
	"Subject: Lunch",
	"",
	"Lunch this week?",  # 20
]


class TestLabels(unittest.TestCase):
	def test_englishAndFrenchAndGerman(self):
		self.assertEqual(labelOf("From: Jane"), ("from", "Jane"))
		self.assertEqual(labelOf("Envoyé : mardi 21 octobre"), ("sent", "mardi 21 octobre"))
		self.assertEqual(labelOf("Betreff: Mittagessen"), ("subject", "Mittagessen"))

	def test_aLineThatIsNotAHeader(self):
		self.assertEqual(labelOf("Note: bring the slides")[0], None)
		self.assertEqual(labelOf("Lunch this week?")[0], None)


class TestFindingAHeader(unittest.TestCase):
	def test_anOutlookHeaderBlock(self):
		self.assertEqual(
			messageAt(OUTLOOK_CHAIN, 4),
			Message(4, 9, "Jane Doe <jane@example.com>", "Tuesday, October 21, 2026 3:15 PM"),
		)

	def test_aFromLineInASignatureIsNot(self):
		self.assertIsNone(messageAt(OUTLOOK_CHAIN, 12))

	def test_aBlockWithoutWhenItWasSentIsNot(self):
		self.assertIsNone(messageAt(["From: Jane", "To: Travis", "", "Hi"], 0))

	def test_anOriginalMessageSeparator(self):
		chain = ["-----Original Message-----", "From: Jane", "Sent: Monday", "To: Travis", "", "Hi"]
		self.assertEqual(messageAt(chain, 0), Message(0, 5, "Jane", "Monday"))

	def test_aGmailForward(self):
		chain = [
			"---------- Forwarded message ---------",
			"From: Jane <jane@example.com>",
			"Date: Mon, Oct 20, 2026 at 9:02 AM",
			"Subject: Lunch",
			"To: <travis@example.com>",
			"",
			"Lunch this week?",
		]
		self.assertEqual(messageAt(chain, 0).content, 6)

	def test_anAppleMailForward(self):
		self.assertEqual(messageAt(["Begin forwarded message:", "", "Hello"], 0).content, 2)

	def test_aGmailReplyLine(self):
		chain = ["On Mon, Oct 20, 2026 at 9:02 AM Jane Doe <jane@example.com> wrote:", "", "Lunch?"]
		self.assertEqual(
			messageAt(chain, 0),
			Message(0, 2, "Jane Doe <jane@example.com>", "Mon, Oct 20, 2026 at 9:02 AM"),
		)

	def test_aReplyLineWrappedOverTwoParagraphs(self):
		chain = ["On Mon, Oct 20, 2026 at 9:02 AM Jane Doe <", "jane@example.com> wrote:", "Lunch?"]
		self.assertEqual(
			messageAt(chain, 0),
			Message(0, 2, "Jane Doe <jane@example.com>", "Mon, Oct 20, 2026 at 9:02 AM"),
		)

	def test_anAppleMailReplyLineSaysWhoAndWhen(self):
		chain = ["On Oct 27, 2026, at 2:13 PM, Person A <a@example.com> wrote:", "Hi"]
		self.assertEqual(
			messageAt(chain, 0), Message(0, 1, "Person A <a@example.com>", "Oct 27, 2026, at 2:13 PM")
		)

	def test_aReplyLineThatCannotBeSplitIsGivenWhole(self):
		line = "Le 20 oct. 2026 à 09:02, Jeanne <j@example.com> a écrit :"
		self.assertEqual(messageAt([line, "Bonjour"], 0), Message(0, 1, header=line))

	def test_aFrenchHeaderBlock(self):
		chain = [
			"De : Jeanne",
			"Envoyé : mardi 21 octobre 2026 15:15",
			"À : Travis",
			"Objet : Déjeuner",
			"",
			"Mardi me va.",
		]
		self.assertEqual(messageAt(chain, 0), Message(0, 5, "Jeanne", "mardi 21 octobre 2026 15:15"))

	def test_aGermanReplyLine(self):
		chain = ["Am 20.10.2026 um 09:02 schrieb Jana <jana@example.com>:", "Mittagessen?"]
		self.assertEqual(messageAt(chain, 0).content, 1)


class TestMovingThroughAChain(unittest.TestCase):
	def test_fromTheTopToTheFirstQuotedMessage(self):
		self.assertEqual(nextMessage(OUTLOOK_CHAIN, 0).content, 9)

	def test_onToTheNext(self):
		self.assertEqual(nextMessage(OUTLOOK_CHAIN, 9).content, 20)

	def test_fromInsideAHeaderToTheNext(self):
		self.assertEqual(nextMessage(OUTLOOK_CHAIN, 5).content, 20)

	def test_theLastHasNoNext(self):
		self.assertIsNone(nextMessage(OUTLOOK_CHAIN, 20))

	def test_backToThePrevious(self):
		self.assertEqual(previousMessage(OUTLOOK_CHAIN, 20, fromTheTop=True).content, 9)

	def test_backToTheTop(self):
		self.assertEqual(previousMessage(OUTLOOK_CHAIN, 10, fromTheTop=True), Message(0, 0))

	def test_theTopHasNoPrevious(self):
		self.assertIsNone(previousMessage(OUTLOOK_CHAIN, 2, fromTheTop=True))

	def test_theTopIsNotOfferedWhenTheParagraphsDoNotReachIt(self):
		self.assertIsNone(previousMessage(OUTLOOK_CHAIN[4:], 6, fromTheTop=False))

	def test_fromAHeaderBackToTheMessageBeforeIt(self):
		self.assertEqual(previousMessage(OUTLOOK_CHAIN, 15, fromTheTop=True).content, 9)

	def test_blankLinesAtTheTopAreSkipped(self):
		self.assertEqual(previousMessage(["", "", *OUTLOOK_CHAIN], 12, fromTheTop=True), Message(0, 2))


BREAK = "\x0b"
"""Word's manual line break, which Outlook ends each line of a reply's header with."""

NBSP = "\xa0"

# Shaped like a real chain from Outlook: a reply over a line of underscores, a reply quoted with
# "On ... wrote:", a forward, and the forwarded message. The names and addresses are made up.
MIXED_CHAIN = [
	"Yes, my thoughts too.",  # 0
	"I will start drafting a response.",
	"Person A",
	BREAK.join(["Person A", "Director", "Example Org", "Direct: (555) 555-0100", "Email: a@example.com"]),
	"Please consider the environment before printing this email.",
	"________________________________________",  # 5
	BREAK.join(
		[
			"From: Person B <b@example.com>",
			"Sent: Tuesday, October 27, 2026 2:31 PM",
			"To: Person A <a@example.com>",
			"Cc: Person C <c@example.com>",
			"Subject: Re: Proposal request ",
		],
	),
	NBSP,
	"I think it looks great!",  # 8
	"It looks promising.",
	BREAK.join(["Thanks,", "Person B"]),
	BREAK.join(["Person B", "Example Org"]),
	"On Oct 27, 2026, at 2:13 PM, Person A <a@example.com> wrote:",  # 12
	NBSP,
	"This looks exciting to me.",  # 14
	"________________________________________",  # 15
	BREAK.join(
		[
			"From: Person B <b@example.com>",
			"Sent: Monday, October 26, 2026 6:02 PM",
			"To: Person A <a@example.com>; Person C <c@example.com>",
			"Subject: Fwd: Proposal request ",
		],
	),
	NBSP,
	"Hi both,",  # 18
	"Person B",
	"Begin forwarded message:",  # 20
	BREAK.join(
		[
			'From: "Person D" <d@example.org>',
			"Date: October 26, 2026 at 12:20:55 PM EDT",
			"To: Person B <b@example.com>",
			"Subject: Proposal request",
		],
	),
	f"{NBSP}{BREAK}{NBSP}",
	"Hi Person B,",  # 23
]


class TestAHeaderInOneParagraph(unittest.TestCase):
	"""Outlook writes a reply's header as one paragraph, a line break after each field."""

	def test_eachMessageInTurn(self):
		contents = []
		current = 0
		while True:
			found = nextMessage(MIXED_CHAIN, current)
			if found is None:
				break
			contents.append(found.content)
			current = found.content
		self.assertEqual(contents, [8, 14, 18, 23])

	def test_theHeaderUnderTheUnderscoresSaysWhoAndWhen(self):
		found = nextMessage(MIXED_CHAIN, 0)
		self.assertEqual(
			found, Message(5, 8, "Person B <b@example.com>", "Tuesday, October 27, 2026 2:31 PM")
		)

	def test_theForwardedMessagesHeader(self):
		found = nextMessage(MIXED_CHAIN, 18)
		self.assertEqual(
			found, Message(20, 23, '"Person D" <d@example.org>', "October 26, 2026 at 12:20:55 PM EDT")
		)

	def test_eachMessageBackInTurn(self):
		contents = []
		current = len(MIXED_CHAIN) - 1
		while True:
			found = previousMessage(MIXED_CHAIN, current, fromTheTop=True)
			if found is None:
				break
			contents.append(found.content)
			current = found.content
		self.assertEqual(contents, [18, 14, 8, 0])

	def test_aSignatureLineWithAColonIsNotAHeader(self):
		self.assertIsNone(messageAt(MIXED_CHAIN, 3))

	def test_underscoresWithNoHeaderUnderThemAreNotOne(self):
		self.assertIsNone(messageAt(["________________________________________", "Just a line."], 0))

	def test_aSeparatorAndItsHeaderInOneParagraph(self):
		chain = [
			BREAK.join(["Begin forwarded message:", "From: Person D", "Date: Monday", "Subject: Hi"]),
			"Hello",
		]
		self.assertEqual(messageAt(chain, 0), Message(0, 1, "Person D", "Monday"))

	def test_underscoresAndTheirHeaderInOneParagraph(self):
		chain = [BREAK.join(["_" * 32, "From: Person B", "Sent: Monday", "To: Person A"]), NBSP, "Hello"]
		self.assertEqual(messageAt(chain, 0), Message(0, 2, "Person B", "Monday"))


if __name__ == "__main__":
	unittest.main()
