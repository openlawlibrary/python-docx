"""Custom element classes related to document settings."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.section import CT_FtnEdnProps
    from docx.oxml.shared import CT_OnOff


class CT_Settings(BaseOxmlElement):
    """`w:settings` element, root element for the settings part."""

    get_or_add_evenAndOddHeaders: Callable[[], CT_OnOff]
    get_or_add_footnotePr: Callable[[], CT_FtnEdnProps]
    get_or_add_endnotePr: Callable[[], CT_FtnEdnProps]
    _remove_evenAndOddHeaders: Callable[[], None]

    _tag_seq = (
        "w:writeProtection",
        "w:view",
        "w:zoom",
        "w:removePersonalInformation",
        "w:removeDateAndTime",
        "w:doNotDisplayPageBoundaries",
        "w:displayBackgroundShape",
        "w:printPostScriptOverText",
        "w:printFractionalCharacterWidth",
        "w:printFormsData",
        "w:embedTrueTypeFonts",
        "w:embedSystemFonts",
        "w:saveSubsetFonts",
        "w:saveFormsData",
        "w:mirrorMargins",
        "w:alignBordersAndEdges",
        "w:bordersDoNotSurroundHeader",
        "w:bordersDoNotSurroundFooter",
        "w:gutterAtTop",
        "w:hideSpellingErrors",
        "w:hideGrammaticalErrors",
        "w:activeWritingStyle",
        "w:proofState",
        "w:formsDesign",
        "w:attachedTemplate",
        "w:linkStyles",
        "w:stylePaneFormatFilter",
        "w:stylePaneSortMethod",
        "w:documentType",
        "w:mailMerge",
        "w:revisionView",
        "w:trackRevisions",
        "w:doNotTrackMoves",
        "w:doNotTrackFormatting",
        "w:documentProtection",
        "w:autoFormatOverride",
        "w:styleLockTheme",
        "w:styleLockQFSet",
        "w:defaultTabStop",
        "w:autoHyphenation",
        "w:consecutiveHyphenLimit",
        "w:hyphenationZone",
        "w:doNotHyphenateCaps",
        "w:showEnvelope",
        "w:summaryLength",
        "w:clickAndTypeStyle",
        "w:defaultTableStyle",
        "w:evenAndOddHeaders",
        "w:bookFoldRevPrinting",
        "w:bookFoldPrinting",
        "w:bookFoldPrintingSheets",
        "w:drawingGridHorizontalSpacing",
        "w:drawingGridVerticalSpacing",
        "w:displayHorizontalDrawingGridEvery",
        "w:displayVerticalDrawingGridEvery",
        "w:doNotUseMarginsForDrawingGridOrigin",
        "w:drawingGridHorizontalOrigin",
        "w:drawingGridVerticalOrigin",
        "w:doNotShadeFormData",
        "w:noPunctuationKerning",
        "w:characterSpacingControl",
        "w:printTwoOnOne",
        "w:strictFirstAndLastChars",
        "w:noLineBreaksAfter",
        "w:noLineBreaksBefore",
        "w:savePreviewPicture",
        "w:doNotValidateAgainstSchema",
        "w:saveInvalidXml",
        "w:ignoreMixedContent",
        "w:alwaysShowPlaceholderText",
        "w:doNotDemarcateInvalidXml",
        "w:saveXmlDataOnly",
        "w:useXSLTWhenSaving",
        "w:saveThroughXslt",
        "w:showXMLTags",
        "w:alwaysMergeEmptyNamespace",
        "w:updateFields",
        "w:hdrShapeDefaults",
        "w:footnotePr",
        "w:endnotePr",
        "w:compat",
        "w:docVars",
        "w:rsids",
        "m:mathPr",
        "w:attachedSchema",
        "w:themeFontLang",
        "w:clrSchemeMapping",
        "w:doNotIncludeSubdocsInStats",
        "w:doNotAutoCompressPictures",
        "w:forceUpgrade",
        "w:captions",
        "w:readModeInkLockDown",
        "w:smartTagType",
        "sl:schemaLibrary",
        "w:shapeDefaults",
        "w:doNotEmbedSmartTags",
        "w:decimalSymbol",
        "w:listSeparator",
    )
    evenAndOddHeaders: CT_OnOff | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:evenAndOddHeaders", successors=_tag_seq[48:]
    )
    footnotePr: CT_FtnEdnProps | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:footnotePr", successors=_tag_seq[79:]
    )
    endnotePr: CT_FtnEdnProps | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:endnotePr", successors=_tag_seq[80:]
    )
    del _tag_seq

    @property
    def endnote_position(self) -> str | None:
        """Document-level default endnote position: value of
        `./w:endnotePr/w:pos/@w:val`, or |None| if not present."""
        endnotePr = self.endnotePr
        return None if endnotePr is None else endnotePr.pos_val

    @endnote_position.setter
    def endnote_position(self, value: str | None):
        if value is None:
            return
        self.get_or_add_endnotePr().pos_val = value

    @property
    def evenAndOddHeaders_val(self) -> bool:
        """Value of `w:evenAndOddHeaders/@w:val` or |None| if not present."""
        evenAndOddHeaders = self.evenAndOddHeaders
        if evenAndOddHeaders is None:
            return False
        return evenAndOddHeaders.val

    @evenAndOddHeaders_val.setter
    def evenAndOddHeaders_val(self, value: bool | None):
        if value is None or value is False:
            self._remove_evenAndOddHeaders()
            return

        self.get_or_add_evenAndOddHeaders().val = value

    @property
    def footnote_position(self) -> str | None:
        """Document-level default footnote position: value of
        `./w:footnotePr/w:pos/@w:val`, or |None| if not present."""
        footnotePr = self.footnotePr
        return None if footnotePr is None else footnotePr.pos_val

    @footnote_position.setter
    def footnote_position(self, value: str | None):
        if value is None:
            return
        self.get_or_add_footnotePr().pos_val = value
