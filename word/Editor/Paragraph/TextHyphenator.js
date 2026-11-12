/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window)
{
	let hyphenator = null;
	function getInstance()
	{
		if (!hyphenator)
			hyphenator = new TextHyphenator();
		
		return hyphenator;
	}
	
	function getWaitingLangCallback(lang, hyphenator)
	{
		return function()
		{
			hyphenator.onLoadLang(lang);
		};
	}
	
	const DEFAULT_LANG = lcid_enUS;
	
	/**
	 * Класс для автоматической расстановки переносов в тексте
	 * @constructor
	 */
	function TextHyphenator()
	{
		this.word     = false;
		this.fontSlot = fontslot_Unknown;
		this.lang     = DEFAULT_LANG;
		this.buffer   = [];
		
		this.hyphenateCaps = true;
		
		this.document     = null;
		this.paragraph    = null;
		this.waitingLangs = {};
	}
	TextHyphenator.hyphenate = function(paragraph)
	{
		let hyphenator = getInstance();
		hyphenator.hyphenateParagraph(paragraph);
	};
	TextHyphenator.reset = function()
	{
		if (!hyphenator)
			return;
		
		hyphenator.reset();
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	TextHyphenator.prototype.hyphenateParagraph = function(paragraph)
	{
		if (!this.document)
			this.document = paragraph.GetLogicDocument();
		
		this.paragraph = paragraph;
		this.checkHyphenateCaps(paragraph);
		
		let self = this;
		paragraph.CheckRunContent(function(run, startPos, endPos)
		{
			self.hyphenateRun(run, startPos, endPos);
		});
		this.flushWord();
	};
	TextHyphenator.prototype.reset = function()
	{
		this.word     = false;
		this.fontSlot = fontslot_Unknown;
		this.lang     = DEFAULT_LANG;
		this.resetBuffer();

		this.hyphenateCaps = true;
		
		this.document     = null;
		this.paragraph    = null;
		this.waitingLangs = {};
	};
	TextHyphenator.prototype.hyphenateRun = function(run, startPos, endPos)
	{
		for (let pos = startPos; pos < endPos; ++pos)
		{
			let item = run.GetElement(pos);
			if (!item.IsText())
			{
				this.flushWord();
			}
			else if (item.IsNBSP() || item.IsPunctuation())
			{
				this.flushWord();
			}
			else
			{
				if (!this.word)
					this.updateLang(run, item.GetFontSlot(run.Get_CompiledPr(false)));
				
				this.appendToWord(item);
				
				if (item.IsSpaceAfter())
					this.flushWord();
			}
		}
	};
	TextHyphenator.prototype.resetBuffer = function()
	{
		this.buffer.length = 0;
		AscHyphenation.clear();
	};
	TextHyphenator.prototype.updateLang = function(run, fontSlot)
	{
		let textPr = run.Get_CompiledPr(false);
		let lang;
		switch (fontSlot)
		{
			case fontslot_EastAsia:
				lang = textPr.Lang.EastAsia;
				break;
			case fontslot_CS:
				lang = textPr.Lang.Bidi;
				break;
			case fontslot_HAnsi:
			case fontslot_ASCII:
			default:
				lang = textPr.Lang.Val;
				break;
		}
		
		if (textPr.CS)
		{
			lang     = textPr.Lang.Bidi;
			fontSlot = fontslot_CS;
		}
		
		this.lang     = lang;
		this.fontSlot = fontSlot;
	};
	TextHyphenator.prototype.appendToWord = function(textItem)
	{
		this.word = true;
		this.buffer.push(textItem);
		AscHyphenation.addCodePoint(String.fromCodePoint(textItem.GetCodePoint()).toLowerCase().codePointAt(0));
		textItem.SetHyphenAfter(false);
	};
	TextHyphenator.prototype.flushWord = function()
	{
		if (!this.word)
			return;
		
		this.word = false;
		
		if (!this.isHyphenateCaps() && this.isAllCaps())
			return this.resetBuffer();
		
		if (!this.checkLangAvailability())
			return this.resetBuffer();
		
		let result = AscHyphenation.hyphenate();
		for (let i = 0, len = result.length; i < len; ++i)
		{
			let pos = result[i] - 1;
			if (pos < 0 || pos >= this.buffer.length - 1)
				continue;
			
			this.buffer[pos].SetHyphenAfter(true);
		}
		
		this.resetBuffer();
	};
	TextHyphenator.prototype.isAllCaps = function()
	{
		for (let i = 0, len = this.buffer.length; i < len; ++i)
		{
			let char = String.fromCodePoint(this.buffer[i].GetCodePoint());
			if (char.toUpperCase() !== char)
				return false;
		}
		
		return true;
	};
	TextHyphenator.prototype.checkHyphenateCaps = function(paragraph)
	{
		this.hyphenateCaps = true;
		
		let logicDocument = paragraph.GetLogicDocument();
		if (logicDocument && logicDocument.IsDocumentEditor())
			this.hyphenateCaps = logicDocument.GetDocumentSettings().isHyphenateCaps();
	};
	TextHyphenator.prototype.isHyphenateCaps = function()
	{
		return this.hyphenateCaps;
	};
	TextHyphenator.prototype.checkLangAvailability = function()
	{
		if (this.waitingLangs[this.lang])
		{
			this.waitingLangs[this.lang][this.paragraph.GetId()] = this.paragraph;
			return false;
		}
		
		if (AscHyphenation.setLang(this.lang, getWaitingLangCallback(this.lang, this)))
			return true;
		
		if (!this.waitingLangs[this.lang])
			this.waitingLangs[this.lang] = {};
		
		this.waitingLangs[this.lang][this.paragraph.GetId()] = this.paragraph;
		return false;
	};
	TextHyphenator.prototype.onLoadLang = function(lang)
	{
		if (!this.waitingLangs[lang])
			return;
		
		let paragraphs = [];
		for (let paraId in this.waitingLangs[lang])
		{
			this.waitingLangs[lang][paraId].NeedHyphenateText();
			paragraphs.push(this.waitingLangs[lang][paraId]);
		}
		delete this.waitingLangs[lang];
		
		if (this.document)
		{
			let history    = this.document.GetHistory();
			let recalcData = history.getRecalcDataByElements(paragraphs);
			this.document.RecalculateWithParams(recalcData);
			this.document.NeedUpdateTarget = false;
		}
	};
	
	/**
	 * #371: Thai, Lao, Khmer and Myanmar are written without spaces between words, so a run of such text is one
	 * "word" for the line breaker and is cut at whatever letter reaches the margin. When the option is on, the
	 * word boundaries found by the browser's dictionary-based Intl.Segmenter become line-break opportunities
	 * (as in Word and LibreOffice). Off by default: it changes the layout of existing documents.
	 */
	let complexScriptLineBreak = false;
	let complexScriptLineBreakId = 0;
	let wordSegmenter = null;
	
	function isNoSpaceScript(codePoint)
	{
		return ((0x0E00 <= codePoint && codePoint <= 0x0EFF)     // Thai, Lao
			|| (0x1000 <= codePoint && codePoint <= 0x109F)      // Myanmar
			|| (0x1780 <= codePoint && codePoint <= 0x17FF)      // Khmer
			|| (0x19E0 <= codePoint && codePoint <= 0x19FF)      // Khmer symbols
			|| (0xAA60 <= codePoint && codePoint <= 0xAA7F));    // Myanmar extended-A
	}
	function getWordSegmenter()
	{
		// bracket access: the property names are not in the compiler externs
		if (!wordSegmenter && typeof Intl !== "undefined" && Intl["Segmenter"])
		{
			try
			{
				wordSegmenter = new Intl["Segmenter"]("th", {"granularity" : "word"});
			}
			catch (e)
			{
				wordSegmenter = null;
			}
		}
		return wordSegmenter;
	}
	const ComplexScriptLineBreaker = {
		setEnabled : function(isOn)
		{
			isOn = !!isOn;
			if (isOn === complexScriptLineBreak)
				return false;
			
			complexScriptLineBreak = isOn;
			++complexScriptLineBreakId;
			return true;
		},
		isEnabled : function()
		{
			return complexScriptLineBreak;
		},
		getId : function()
		{
			return complexScriptLineBreakId;
		},
		/**
		 * @param paragraph
		 * @returns {?Set} text items after which a line may break, or null
		 */
		getBreaks : function(paragraph)
		{
			if (!complexScriptLineBreak)
				return null;
			
			let text = "", items = [], hasScript = false;
			paragraph.CheckRunContent(function(run, startPos, endPos)
			{
				for (let pos = startPos; pos < endPos; ++pos)
				{
					let item = run.GetElement(pos);
					if (item.IsText())
					{
						let codePoint = item.GetCodePoint();
						let str = String.fromCodePoint(codePoint);
						if (isNoSpaceScript(codePoint))
							hasScript = true;
						
						text += str;
						for (let i = 0; i < str.length; ++i)
							items.push(item);
					}
					else
					{
						text += " ";
						items.push(null);
					}
				}
			});
			
			let segmenter = hasScript ? getWordSegmenter() : null;
			if (!segmenter)
				return null;
			
			let breaks = new Set();
			let segments = segmenter["segment"](text)[Symbol.iterator]();
			for (let next = segments.next(); !next.done; next = segments.next())
			{
				let index = next.value["index"];
				if (index <= 0 || index >= items.length)
					continue;
				
				let before = items[index - 1], after = items[index];
				if (before && after && before !== after
					&& isNoSpaceScript(before.GetCodePoint())
					&& isNoSpaceScript(after.GetCodePoint()))
					breaks.add(before);
			}
			return breaks.size ? breaks : null;
		}
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].TextHyphenator = TextHyphenator;
	window['AscWord'].ComplexScriptLineBreaker = ComplexScriptLineBreaker;
	
})(window);

