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

// Import
var c_oAscError = Asc.c_oAscError;

/////////////////////////////////////////////////////////
//////////////        OPEN       ////////////////////////
/////////////////////////////////////////////////////////

(/**
 * @param {jQuery} $
 * @param {Window} window
 * @param {undefined} undefined
 */
	function($, window, undefined) {

	var asc = window["Asc"];
	var spreadsheet_api = asc['spreadsheet_api'];

	spreadsheet_api.prototype.asc_addImageDrawingObject = function(urls, imgProp, token)
	{
		var ws = this.wb.getWorksheet();
		if (ws) 
		{
			var _urls = urls.map(function(currentValue) {
				var localUrl =  window["AscDesktopEditor"]["LocalFileGetImageUrl"](currentValue);
				return AscCommon.g_oDocumentUrls.getImageUrl(localUrl);
			});
			ws.objectRender.addImageDrawingObject(_urls, null);
		}
	};
	spreadsheet_api.prototype.asc_showImageFileDialog = spreadsheet_api.prototype.asc_addImage = function(obj)
	{
		var t = this;
		window["AscDesktopEditor"]["OpenFilenameDialog"]("images", true, function(files) {
			// #617: several images can be picked at once, like the online editors' file chooser
			var urls = AscCommon.getDesktopImageUrls(files);
			if (!urls.length)
				return;

			t._addImageUrl(urls, obj);
		});
	};
	spreadsheet_api.prototype.AddImageUrlAction = function(url)
	{
		this._addImageUrl([url]);
	};
	spreadsheet_api.prototype.asc_setAdvancedOptions = function(idOption, option)
	{
		if (asc.c_oAscAdvancedOptionsID.CSV === idOption) {
			var delimiter = option.asc_getDelimiter();
			var delimiterChar = option.asc_getDelimiterChar();
			var _param = "";
			_param += ("<m_nCsvTxtEncoding>" + option.asc_getCodePage() + "</m_nCsvTxtEncoding>");
			if (null != delimiter) {
				_param += ("<m_nCsvDelimiter>" + delimiter + "</m_nCsvDelimiter>");
			}
			if (null != delimiterChar) {
				_param += ("<m_nCsvDelimiterChar>" + delimiterChar + "</m_nCsvDelimiterChar>");
			}

			window["AscDesktopEditor"]["SetAdvancedOptions"](_param);
		}
		else if (asc.c_oAscAdvancedOptionsID.DRM === idOption) {
			var _param = "";
			_param += ("<m_sPassword>" + AscCommon.CopyPasteCorrectString(option.asc_getPassword()) + "</m_sPassword>");
			this.currentPassword = option.asc_getPassword();
			window["AscDesktopEditor"]["SetAdvancedOptions"](_param);
		}
	};
	/////////////////////////////////////////////////////////
	////////////////        SAVE       //////////////////////
	/////////////////////////////////////////////////////////
	spreadsheet_api.prototype.onUpdateDocumentModified = function(bIsModified)
	{
		// Update only after saving is complete
		if (this.canSave) {
			this.handlers.trigger("asc_onDocumentModifiedChanged", bIsModified);
			this._onUpdateDocumentCanSave();

			if (undefined !== window["AscDesktopEditor"]) {
				window["AscDesktopEditor"]["onDocumentModifiedChanged"](AscCommon.History ? AscCommon.History.Have_Changes(undefined, true) : bIsModified);
			}
		}
	};

	spreadsheet_api.prototype._saveLocalCheck = function()
	{
		return this._saveCheck();
	};
	spreadsheet_api.prototype._isSaveDroppingSheets = function()
	{
		// the file is saved back in place when it has a path; csv/tsv hold one sheet
		var sPath = window["AscDesktopEditor"]["LocalFileGetSourcePath"] ? window["AscDesktopEditor"]["LocalFileGetSourcePath"]() : "";
		return !!(this.wbModel && this.wbModel.getWorksheetCount() > 1 && /\.(csv|tsv)$/i.test(sPath || ""));
	};
	spreadsheet_api.prototype.asc_Save = function (isNoUserSave, isSaveAs, isResaveAttack, options)
	{
		if (this.isFrameEditor() || AscCommon.c_oAscAdvancedOptionsAction.None !== this.advancedOptionsAction)
			return;

		//is cell editor active - try to close
		this.asc_closeCellEditor();

		var t = this;

		// #1490: a plain Save of a file opened from .csv/.tsv writes that format again, and x2t
		// keeps only the active sheet - the other sheets were dropped without a word (Save As
		// CSV already warns). Ask first: "save" keeps the format and saves the active sheet,
		// anything else opens Save As so the workbook can go to a format that keeps it. Both
		// end in the host's own save, so a close waiting for this save is never left hanging.
		if (true !== isNoUserSave && true !== isSaveAs && true !== this.isCsvSheetsSaveConfirmed &&
			true === this.canSave && this._saveLocalCheck() && this._isSaveDroppingSheets() &&
			this.handlers.handlers && this.handlers.handlers["asc_onConfirmAction"])
		{
			this.handlers.trigger("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmSaveCsvSheets, function (can) {
				t.isCsvSheetsSaveConfirmed = true;
				t.asc_Save(isNoUserSave, !can, isResaveAttack, options);
				t.isCsvSheetsSaveConfirmed = false;
			});
			return;
		}

		if (true !== isNoUserSave)
			this.IsUserSave = true;

		if (this.IsUserSave)
		{
			this.LastUserSavedIndex = AscCommon.History.UserSavedIndex;
		}

		if (true === this.canSave && this._saveLocalCheck())
		{
			var _isNaturalSave = this.IsUserSave;
			this.canSave = false;
			this.CoAuthoringApi.askSaveChanges(function(e){t._onSaveCallback(e);});

			if (this.CoAuthoringApi.onUnSaveLock)
				this.CoAuthoringApi.onUnSaveLock();

			if (_isNaturalSave === true)
				window["DesktopOfflineAppDocumentStartSave"](isSaveAs, undefined, undefined, undefined, options);
		}
	};
    spreadsheet_api.prototype.asc_DownloadAsNatural = spreadsheet_api.prototype.asc_DownloadAs;
	spreadsheet_api.prototype.asc_DownloadAs = function(options)
	{
        if (options && options.isNaturalDownload)
            return this.asc_DownloadAsNatural(options);
		this.asc_Save(false, true, undefined, options);
	};
	spreadsheet_api.prototype.asc_isOffline = function()
	{
		return true;
	};

	spreadsheet_api.prototype["asc_setAdvancedOptions"] = spreadsheet_api.prototype.asc_setAdvancedOptions;
	spreadsheet_api.prototype["asc_addImageDrawingObject"] = spreadsheet_api.prototype.asc_addImageDrawingObject;
	spreadsheet_api.prototype["asc_showImageFileDialog"] = spreadsheet_api.prototype.asc_showImageFileDialog;
	spreadsheet_api.prototype["asc_Save"] = spreadsheet_api.prototype.asc_Save;
	spreadsheet_api.prototype["asc_DownloadAs"] = spreadsheet_api.prototype.asc_DownloadAs;
	spreadsheet_api.prototype["asc_isOffline"] = spreadsheet_api.prototype.asc_isOffline;
	spreadsheet_api.prototype["asc_addImage"] = spreadsheet_api.prototype.asc_addImage;

	spreadsheet_api.prototype["getAdditionalSaveParams"] = function()
	{
		var printOptionsObj = asc["editor"].getPrintOptionsJson();
		printOptionsObj["documentLayout"] = { "openedAt" : asc["editor"].openedAt };
		printOptionsObj["locale"] = asc["editor"].asc_getLocale();
		printOptionsObj["translate"] = AscCommon.translateManager.mapTranslate;

		// Which sheet the user is on is pure view state: Workbook.prototype.setActive
		// writes nActive with no History entry, so it never enters the change stream.
		// A desktop save ships *changes*, which x2t applies to the Editor.bin written
		// when the document was opened, so the activeTab stored there is whatever it
		// was at open time - and that is the sheet x2t exports to CSV. Ship the real
		// one in the save parameters instead. ONLYOFFICE/DesktopEditors#1839.
		var wbModel = asc["editor"].wbModel;
		if (wbModel)
			printOptionsObj["activeSheet"] = wbModel.getActive();

		// Where inside a sheet the user is looking is view state too, and worse off
		// than the active sheet. Nothing copies the live scroll position into
		// Worksheet.sheetViews[0].topLeftCell while the user scrolls - the two
		// WorksheetView calls to model.updateTopLeftCell are commented out - and the
		// one path that does, WorkbookView.executeWithCurrentTopLeftCell, runs only
		// around a full-binary write (cell/api.js asc_DownloadAs), which a desktop
		// save does not do: it ships *changes*, which x2t applies to the Editor.bin
		// written at open time. So the saved file carries the scroll position from
		// when it was opened, and reopening lands on A1. Ship the live top-left cell
		// of every sheet that has a view, as "<sheet index>:<A1-style ref>" pairs.
		// ONLYOFFICE/DesktopEditors#1868.
		var wbView = asc["editor"].wb;
		if (wbView && wbView.wsViews) {
			var aTopLeftCells = [];
			for (var nSheet in wbView.wsViews) {
				var oWS = wbView.wsViews[nSheet];
				if (!oWS)
					continue;
				var oTopLeftCell = oWS.getCurrentTopLeftCell();
				// null is the model's own way of spelling A1: see
				// Worksheet.prototype.generateTopLeftCellFromRange, which returns it
				// both for the origin and for the frozen-pane case it cannot express,
				// and executeWithCurrentTopLeftCell, which stores exactly this value
				// before a full-binary write. Send it explicitly, so that scrolling
				// back to the top also clears a stale stored position.
				// Formatted from the range rather than with getName(), because
				// getName() honours g_R1C1Mode and the file format wants A1 style.
				var sRef = "A1";
				if (oTopLeftCell) {
					sRef = AscCommon.g_oCellAddressUtils.colnumToColstr(oTopLeftCell.c1 + 1) +
						(oTopLeftCell.r1 + 1);
				}
				aTopLeftCells.push(nSheet + ":" + sRef);
			}
			if (aTopLeftCells.length)
				printOptionsObj["topLeftCells"] = aTopLeftCells.join(";");

			// The caret is view state in the same way: moving it writes nothing to the
			// change stream, so the saved file kept the selection from open time. Ship each
			// viewed sheet's active cell and the range holding it, as
			// "<sheet index>:<active cell>:<range>" (A1 style, ';' between sheets).
			var aActiveCells = [];
			var fA1 = function (row, col) {
				return AscCommon.g_oCellAddressUtils.colnumToColstr(col + 1) + (row + 1);
			};
			for (var nSheetSel in wbView.wsViews) {
				var oWSSel = wbView.wsViews[nSheetSel];
				var oSelection = oWSSel && oWSSel.model && oWSSel.model.selectionRange;
				var oActiveCell = oSelection && oSelection.activeCell;
				var oActiveRange = oSelection && oSelection.getLast();
				if (!oActiveCell || !oActiveRange)
					continue;
				var sRange = fA1(oActiveRange.r1, oActiveRange.c1);
				if (oActiveRange.r1 !== oActiveRange.r2 || oActiveRange.c1 !== oActiveRange.c2)
					sRange += ":" + fA1(oActiveRange.r2, oActiveRange.c2);
				aActiveCells.push(nSheetSel + ":" + fA1(oActiveCell.row, oActiveCell.col) + ":" + sRange);
			}
			if (aActiveCells.length)
				printOptionsObj["activeCells"] = aActiveCells.join(";");
		}

		return printOptionsObj;
	};

	spreadsheet_api.prototype["asc_changeExternalReference"] = function(eR)
	{
		let wb = this.wb;
		if (!wb) {
			return;
		}
		window["AscDesktopEditor"]["OpenFilenameDialog"]("cell", false, function(_file) {
			let file = _file;
			if (Array.isArray(file))
				file = file[0];
			if (!file)
				return;

			//if not saved - put absolute path
			let isSavedFile = window["AscDesktopEditor"]["LocalFileGetSaved"]();
			let relativePath = isSavedFile && window["AscDesktopEditor"]["LocalFileGetRelativePath"](file);

			let obj = {};
			obj["path"] = relativePath;
			obj["filePath"] = file;
			wb.changeExternalReference(eR, obj);
		});
	};

	/////////////////////////////////////////////////////////
	//////////////        CHANGES       /////////////////////
	/////////////////////////////////////////////////////////
	AscCommon.CHistory.prototype.Reset_SavedIndex = function(IsUserSave)
	{
		this.SavedIndex = (null === this.SavedIndex && -1 === this.Index ? null : this.Index);
		if (true === this.Is_UserSaveMode())
		{
			if (true === IsUserSave)
			{
				this.UserSavedIndex = this.Index;
				this.ForceSave      = false;
			}
		}
		else
		{
			this.ForceSave  = false;
		}
	};
	AscCommon.CHistory.prototype.Have_Changes = function(IsNotUserSave, IsNoSavedNoModifyed)
	{
		var checkIndex = (this.Is_UserSaveMode() && !IsNotUserSave) ? this.UserSavedIndex : this.SavedIndex;
		if (-1 === this.Index && null === checkIndex && false === this.ForceSave)
		{
			if (window["AscDesktopEditor"])
			{
				if (0 != window["AscDesktopEditor"]["LocalFileGetOpenChangesCount"]())
					return true;
				if (!window["AscDesktopEditor"]["LocalFileGetSaved"]() && IsNoSavedNoModifyed !== true)
					return true;
			}
			return false;
		}
		return (this.Index != checkIndex || true === this.ForceSave);
	};

	window["DesktopOfflineAppDocumentApplyChanges"] = function(_changes)
	{
		for (var i = 0, l = _changes.length; i < l; ++i)
		{
			asc["editor"].CoAuthoringApi.onSaveChanges(_changes[i], null, true);
		}
	};
	window["DesktopOfflineAppDocumentStartSave"] = function(isSaveAs, password, isForce, docinfo, options)
	{
		window.doadssIsSaveAs = isSaveAs;
		if (true !== isForce && window.g_asc_plugins && AscCommon.EncryptionWorker.isNeedCrypt())
		{
			window.g_asc_plugins.sendToEncryption({ "type" : "generatePassword" });
			return;
		}

		asc["editor"].sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);

		var _param = "";
		if (isSaveAs === true)
			_param += "saveas=true;";
		// #982: "Save Copy as" - the host writes the file but keeps the open document as it is
		if (isSaveAs === true && options && options.isSaveCopy)
			_param += "savecopy=true;";
		// #517: preferred format for the first save of a new document (the host decides if it applies)
		if (asc["editor"].newFileSaveFormat)
			_param += "newformat=" + asc["editor"].newFileSaveFormat + ";";

		// settings for pdf
		var printOptions = "";
		if (isSaveAs)
		{
			try
			{
				var printOptionsObj = asc["editor"]["getAdditionalSaveParams"]();

				if (options && options.advancedOptions)
				{
					let nativeOptions = options.advancedOptions.asc_getNativeOptions();
					if (nativeOptions)
						printOptionsObj["nativeOptions"] = nativeOptions;

					// adjustPrint
					printOptionsObj["adjustOptions"] = {};
					if (null != options.advancedOptions.asc_getStartPageIndex())
						printOptionsObj["adjustOptions"]["startPageIndex"] =  options.advancedOptions.asc_getStartPageIndex();
					if (null != options.advancedOptions.asc_getEndPageIndex())
						printOptionsObj["adjustOptions"]["endPageIndex"] =  options.advancedOptions.asc_getEndPageIndex();
					if (null != options.advancedOptions.asc_getActiveSheetsArray())
						printOptionsObj["adjustOptions"]["activeSheetsArray"] = options.advancedOptions.asc_getActiveSheetsArray();
					if (null != options.advancedOptions.asc_getPrintType())
						printOptionsObj["adjustOptions"]["printType"] = options.advancedOptions.asc_getPrintType();
				}

				printOptions = JSON.stringify(printOptionsObj);
			}
			catch (e)
			{
				printOptions = "";
			}
		}
		else
		{
			// A plain Save needs the view state too: Ctrl+S on a .csv (or picking CSV in
			// the dialog for a new file) goes through here, and without "activeSheet" x2t
			// exports the sheet that was active at open time, not the one on screen.
			// ONLYOFFICE/DesktopEditors#902.
			// A plain Save can also produce a PDF: the first Save of a new workbook (or of
			// one in a format we cannot write) asks where to save, and PDF is offered there.
			// So send the sheets' own page setup and the regional setting, as Save As does.
			// What went here before was the host's no-parameters default (fileconverter.h),
			// fitToWidth/fitToHeight 1, which squeezed every sheet onto one page with tiny
			// text. ONLYOFFICE/DesktopEditors#547.
			try
			{
				var saveParamsObj = asc["editor"]["getAdditionalSaveParams"]();
				var viewParamsObj = { "spreadsheetLayout" : saveParamsObj["spreadsheetLayout"] };
				if (undefined !== saveParamsObj["locale"])
					viewParamsObj["locale"] = saveParamsObj["locale"];
				if (undefined !== saveParamsObj["activeSheet"])
					viewParamsObj["activeSheet"] = saveParamsObj["activeSheet"];
				if (undefined !== saveParamsObj["topLeftCells"])
					viewParamsObj["topLeftCells"] = saveParamsObj["topLeftCells"];
				if (undefined !== saveParamsObj["activeCells"])
					viewParamsObj["activeCells"] = saveParamsObj["activeCells"];

				printOptions = JSON.stringify(viewParamsObj);
			}
			catch (e)
			{
				printOptions = "";
			}
		}

		window["AscDesktopEditor"]["LocalFileSave"](_param, (password === undefined) ? asc["editor"].currentPassword : password, docinfo, (options && options.fileType) ? options.fileType : 0, printOptions);
	};
	window["DesktopOfflineAppDocumentEndSave"] = function(error, hash, password)
	{
		asc["editor"].sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);
		if (0 == error)
			DesktopOfflineUpdateLocalName(asc["editor"]);
		else
			AscCommon.History.UserSavedIndex = asc["editor"].LastUserSavedIndex;

        var _lastUserSavedError = asc["editor"].LastUserSavedIndex;

		asc["editor"].onUpdateDocumentModified(AscCommon.History.Have_Changes());
		asc["editor"].LastUserSavedIndex = undefined;

		if (2 == error)
			asc["editor"].sendEvent("asc_onError", c_oAscError.ID.ConvertationSaveError, c_oAscError.Level.NoCritical);

		if (0 == error)
		{
			if (window.SaveQuestionObjectBeforeSign)
			{
				var _obj = window.SaveQuestionObjectBeforeSign;
				asc["editor"].sendEvent("asc_onSignatureClick", _obj.guid, _obj.width, _obj.height, window["asc_IsVisibleSign"](_obj.guid));
				window.SaveQuestionObjectBeforeSign = null;
			}
		}

		if (hash !== null && hash !== undefined && hash != "")
		{
			if (window.g_asc_plugins && window.g_asc_plugins.isRunnedEncryption())
			{
                asc["editor"]._callbackPluginEndAction = function()
                {
                    this._callbackPluginEndAction = null;
                    window["AscDesktopEditor"]["buildCryptedEnd"](true);
                };
                window.LastUserSavedIndex = _lastUserSavedError;
				window.g_asc_plugins.sendToEncryption({"type": "setPasswordByFile", "hash": hash, "password": password});
			}
		}

		if (0 == error)
			asc["editor"].sendEvent("asc_onDocumentPassword", ("" != asc["editor"].currentPassword) ? true : false);
	};
	window["on_editor_native_message"] = function(sCommand, sParam)
	{
		if (!asc["editor"])
			return;

		if (sCommand == "save")
			asc["editor"].asc_Save();
		else if (sCommand == "saveAs")
			asc["editor"].asc_Save(false, true);
		else if (sCommand == "print")
			asc["editor"].asc_Print();
	};
})(jQuery, window);
