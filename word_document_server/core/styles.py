"""
Style-related functions for Word Document Server.

These are the style entry points the pre-V2 tools call, with the signatures they
have always had.  What changed in J05 is what happens behind them: instead of
python-docx's ``styles.add_style`` -- which never ran here, see
:func:`create_style` -- they delegate to
:mod:`word_document_server.engine.styles`, which writes the ``w:style`` element
itself, with its children in schema order and its references checked.

The ``doc`` argument of every function is either a python-docx ``Document`` or a
:class:`~word_document_server.engine.package.DocxPackage`; the engine accepts
both, and passing the package is what lets a tool edit the style sheet without
rewriting the whole document on the way out.
"""
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH

from word_document_server.core.tables import run_patch
from word_document_server.engine.errors import LocatorError
from word_document_server.engine.styles import create_style as engine_create_style
from word_document_server.engine.styles import style_info

#: ``w:styleId`` and ``w:name`` of Word's nine heading styles.  Word derives
#: neither from the other -- the name is ``heading 1`` and the id ``Heading1`` --
#: so both are stated rather than computed.
_HEADING_IDS = {level: (f"Heading{level}", f"heading {level}") for level in range(1, 10)}

#: Font size in points of each heading level, as this server has always made
#: them: 16 for level 1, 14 for level 2, 12 below.
_HEADING_SIZES = {1: 16.0, 2: 14.0}
_HEADING_SIZE = 12.0

#: ``w:uiPriority`` Word gives every one of its heading styles.
_HEADING_PRIORITY = 9

#: Space above a heading, in twips, and the ``w:outlineLvl`` of level 1.
_HEADING_SPACE_BEFORE = 240

#: The two families a caller may ask for, whatever spelling it uses.
_FAMILIES = {
    WD_STYLE_TYPE.PARAGRAPH: "paragraph",
    WD_STYLE_TYPE.CHARACTER: "character",
    "paragraph": "paragraph",
    "character": "character",
}

#: python-docx alignment enum -> ``w:jc`` value.
_ALIGNMENTS = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "both",
    WD_ALIGN_PARAGRAPH.DISTRIBUTE: "distribute",
}

#: Twips per line of ``w:spacing w:line`` when the rule is ``auto``: a line
#: spacing of 1.5 is ``1.5 * 240``.
_TWIPS_PER_LINE = 240


def ensure_heading_style(doc):
    """
    Ensure Heading 1 through Heading 9 exist, as real heading styles.

    A heading is not a big bold paragraph: it is a paragraph whose style carries
    ``w:outlineLvl``, which is what puts it in the navigation pane, in a table of
    contents and in the document outline.  A style missing that is a paragraph
    that merely looks like a heading, so the styles created here carry the whole
    set Word gives its own: ``basedOn Normal``, ``next Normal``, the outline
    level, ``qFormat`` so Word offers them, and the ``uiPriority`` that sorts
    them in the gallery.

    Nothing is touched when the document already defines the style -- which is
    the usual case, since Word's own template defines all nine.

    Args:
        doc: Document object, or a
            :class:`~word_document_server.engine.package.DocxPackage`
    """
    for level, (style_id, name) in _HEADING_IDS.items():
        try:
            engine_create_style(
                doc,
                {
                    "name": name,
                    "style_id": style_id,
                    "family": "paragraph",
                    "based_on": "Normal",
                    "next": "Normal",
                    "builtin": True,
                    "q_format": True,
                    "ui_priority": _HEADING_PRIORITY,
                    "run_props": {
                        "bold": True,
                        "size_pt": _HEADING_SIZES.get(level, _HEADING_SIZE),
                    },
                    "paragraph_props": {
                        "outline_level": level - 1,
                        "keep_next": True,
                        "keep_lines": True,
                        "spacing": {"before": _HEADING_SPACE_BEFORE, "after": 0},
                    },
                },
            )
        except LocatorError as exc:
            # ``already_exists`` is the expected answer here: "ensure" means
            # create what is missing, not redefine what the document already has,
            # and a document whose Heading 3 is a hand-made style keeps it.
            if exc.code != "already_exists":
                raise


def ensure_table_style(doc):
    """
    Ensure Table Grid style exists in the document.

    Args:
        doc: Document object
    """
    try:
        # Try to access the style to see if it exists
        style = doc.styles['Table Grid']
    except KeyError:
        # If style doesn't exist, we'll handle it at usage time
        pass


def _font_props(font_properties):
    """Translate the legacy ``font_properties`` dict into engine run properties.

    The legacy keys (``size``, ``name``) and the engine's (``size_pt``, ``font``)
    name the same things; the colour goes through the same
    :func:`~word_document_server.core.tables.resolve_color` the rest of the
    server uses, so ``"red"`` means here exactly what it means everywhere else.

    Raises:
        ValueError: if a colour is neither a known name, ``auto`` nor a
            hexadecimal value.  The previous code fell back to black, which is a
            document silently formatted the way nobody asked.
    """
    properties = dict(font_properties or {})
    color = properties.get('color')
    if color is not None and not isinstance(color, str):
        # An RGBColor, or anything else with a string form of six hex digits.
        color = str(color)
    return run_patch(
        bold=properties.get('bold'),
        italic=properties.get('italic'),
        underline=properties.get('underline'),
        color=color,
        font_size=properties.get('size'),
        font_name=properties.get('name'),
    )


def _paragraph_props(paragraph_properties):
    """Translate the legacy ``paragraph_properties`` dict into engine properties.

    ``spacing`` has always meant the line spacing as a multiple of the line
    height, which is what ``w:line`` stores in twentieths of a point under the
    ``auto`` rule.

    Raises:
        ValueError: on an alignment this server cannot express.
    """
    properties = dict(paragraph_properties or {})
    decoded = {}
    if properties.get('alignment') is not None:
        alignment = properties['alignment']
        if isinstance(alignment, str):
            decoded['alignment'] = alignment
        elif alignment in _ALIGNMENTS:
            decoded['alignment'] = _ALIGNMENTS[alignment]
        else:
            raise ValueError(f"unsupported paragraph alignment {alignment!r}")
    if properties.get('spacing') is not None:
        decoded['spacing'] = {
            'line': int(round(float(properties['spacing']) * _TWIPS_PER_LINE)),
            'line_rule': 'auto',
        }
    return decoded


def create_style(doc, style_name, style_type, base_style=None, font_properties=None,
                 paragraph_properties=None):
    """
    Create a new style in the document.

    Until J05 this function created nothing: it guarded the creation with
    ``doc.styles.get_by_id(style_name, WD_STYLE_TYPE.PARAGRAPH)``, and
    ``get_by_id`` never raises -- it returns the default style of the type when
    the id is unknown -- so the ``except`` branch that did the work was never
    reached and the caller was told the style had been created.  It now writes
    the style through :mod:`word_document_server.engine.styles`.

    Args:
        doc: Document object, or a
            :class:`~word_document_server.engine.package.DocxPackage`
        style_name: Name for the new style
        style_type: ``WD_STYLE_TYPE.PARAGRAPH`` or ``WD_STYLE_TYPE.CHARACTER``
            (the strings ``"paragraph"`` and ``"character"`` are accepted too)
        base_style: Optional base style to inherit from, by id or by name
        font_properties: Dictionary of font properties (bold, italic, size, name,
            color)
        paragraph_properties: Dictionary of paragraph properties (alignment,
            spacing)

    Returns:
        The :class:`~word_document_server.engine.styles.StyleInfo` of the style,
        whether it was created here or already defined.

    Raises:
        ValueError: on an unsupported style type, an unreadable colour or an
            alignment this server cannot express.
        LocatorError: ``not_found`` if `base_style` names no style of the
            document.
    """
    family = _FAMILIES.get(style_type)
    if family is None:
        raise ValueError(
            f"a style is created as a paragraph or a character style, got {style_type!r}"
        )
    spec = {
        'name': style_name,
        'family': family,
        'q_format': True,
        'run_props': _font_props(font_properties),
    }
    if base_style:
        spec['based_on'] = base_style
    paragraph_props = _paragraph_props(paragraph_properties)
    if paragraph_props:
        if family != 'paragraph':
            raise ValueError(
                f"style {style_name!r} is a character style; it has no paragraph "
                "properties"
            )
        spec['paragraph_props'] = paragraph_props

    try:
        return engine_create_style(doc, spec)
    except LocatorError as exc:
        if exc.code != 'already_exists':
            raise
        # A style that is already there is not a failure: the caller asked for it
        # to exist, and it does.  What it asked for is *not* applied to it --
        # silently redefining a style the document already uses would reformat
        # every paragraph that names it.
        return style_info(doc, style_name)
