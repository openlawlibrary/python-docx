# encoding: utf-8

"""
The reference marks that Word shows for footnotes and endnotes.

A note has an id, not a number. Word calculates the mark of each reference
from the order of the references in the body and from the note properties
(ECMA-376 part 1, 17.11): the number format, the start number, and when the
numbers restart. The properties of the document (``w:settings``) apply,
unless the section of the reference (``w:sectPr``) sets its own.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

from .ns import qn
from .numbering import CT_Numbering

#: Word's defaults: footnotes 1, 2, 3; endnotes i, ii, iii.
DEFAULT_NOTE_PROPERTIES = {
    'footnote': {'numFmt': 'decimal', 'numStart': 1, 'numRestart': 'continuous'},
    'endnote': {'numFmt': 'lowerRoman', 'numStart': 1, 'numRestart': 'continuous'},
}

NOTE_REFERENCES = {
    qn('w:footnoteReference'): 'footnote',
    qn('w:endnoteReference'): 'endnote',
}

#: A reference in a deleted revision does not show.
DELETED = (qn('w:del'), qn('w:moveFrom'))

CHICAGO_SYMBOLS = '*†‡§'  # *, dagger, double dagger, section


def format_note_number(number, number_format):
    """
    Return *number* as Word shows it in *number_format*. The "chicago"
    format repeats a symbol: *, †, ‡, §, **, ††. A format that is not known
    here gives a decimal number.
    """
    if number_format == 'chicago':
        return CHICAGO_SYMBOLS[(number - 1) % 4] * ((number - 1) // 4 + 1)
    formatter = CT_Numbering.fmt_map.get(number_format, CT_Numbering.fmt_map['decimal'])
    return str(formatter(number))


def note_properties(notePr, inherited):
    """
    Return the numbering properties of a ``<w:footnotePr>`` or
    ``<w:endnotePr>`` element, with the *inherited* value for each property
    that the element does not set.
    """
    properties = dict(inherited)
    if notePr is not None:
        if notePr.numFmt is not None:
            properties['numFmt'] = notePr.numFmt.val
        if notePr.numStart is not None:
            properties['numStart'] = notePr.numStart.val
        if notePr.numRestart is not None:
            properties['numRestart'] = notePr.numRestart.val
    return properties


def custom_mark(reference):
    """
    Return the mark of a reference with ``w:customMarkFollows``: the text
    after the reference in its run, or else the text of the next run.
    """
    text = ''.join(t.text or '' for t in reference.itersiblings(qn('w:t')))
    run = reference.getparent().getnext()
    while not text and run is not None:
        if run.tag == qn('w:r'):
            return ''.join(t.text or '' for t in run.iter(qn('w:t')))
        run = run.getnext()
    return text


def note_reference_marks(body, settings):
    """
    Return ``({footnote id: mark}, {endnote id: mark})`` for the note
    references in *body*, each in the order of the references.

    *settings* is the ``<w:settings>`` element, or |None|. Footnotes and
    endnotes have separate counters. A reference with ``w:customMarkFollows``
    shows its own text and takes no number. A reference in a deleted
    revision does not count. Word restarts the numbers on each page for
    ``numRestart="eachPage"``; that needs the page layout, so here the
    numbers continue. A ``numStart`` in a later section applies only when
    that section restarts the numbers (``eachSect``).
    """
    document_properties = {
        kind: note_properties(
            getattr(settings, '%sPr' % kind, None), DEFAULT_NOTE_PROPERTIES[kind])
        for kind in DEFAULT_NOTE_PROPERTIES
    }

    paragraphs = list(body.iter(qn('w:p')))
    # A section ends at the w:sectPr in its last paragraph; the w:sectPr of
    # the body ends the last section.
    breaks = [p.pPr.sectPr if p.pPr is not None else None for p in paragraphs]
    sections = [sectPr for sectPr in breaks if sectPr is not None]
    sections.append(body.find(qn('w:sectPr')))

    def section_properties(index):
        sectPr = sections[index]
        return {
            kind: note_properties(
                getattr(sectPr, '%sPr' % kind, None), document_properties[kind])
            for kind in DEFAULT_NOTE_PROPERTIES
        }

    section = 0
    properties = section_properties(section)
    counters = {kind: properties[kind]['numStart'] for kind in properties}
    marks = {'footnote': {}, 'endnote': {}}
    for paragraph, section_break in zip(paragraphs, breaks):
        for reference in paragraph.iter(*NOTE_REFERENCES):
            if any(ancestor.tag in DELETED for ancestor in reference.iterancestors()):
                continue
            kind = NOTE_REFERENCES[reference.tag]
            if reference.customMarkFollows:
                mark = custom_mark(reference)
            else:
                mark = format_note_number(counters[kind], properties[kind]['numFmt'])
                counters[kind] += 1
            marks[kind][reference.id] = mark
        if section_break is not None and section + 1 < len(sections):
            section += 1
            properties = section_properties(section)
            for kind in properties:
                if properties[kind]['numRestart'] == 'eachSect':
                    counters[kind] = properties[kind]['numStart']
    return marks['footnote'], marks['endnote']
