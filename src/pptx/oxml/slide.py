"""Slide-related custom element classes, including those for masters."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml import parse_from_template, parse_xml
from pptx.oxml.dml.fill import CT_GradientFillProperties
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
    ZeroOrOneChoice,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.groupshape import CT_GroupShape


class _BaseSlideElement(BaseOxmlElement):
    """Base class for the six slide types, providing common methods."""

    cSld: CT_CommonSlideData

    @property
    def spTree(self) -> CT_GroupShape:
        """Return required `p:cSld/p:spTree` grandchild."""
        return self.cSld.spTree


class CT_Background(BaseOxmlElement):
    """`p:bg` element."""

    _insert_bgPr: Callable[[CT_BackgroundProperties], None]

    # ---these two are actually a choice, not a sequence, but simpler for
    # ---present purposes this way.
    _tag_seq = ("p:bgPr", "p:bgRef")
    bgPr: CT_BackgroundProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bgPr", successors=()
    )
    bgRef = ZeroOrOne("p:bgRef", successors=())
    del _tag_seq

    def add_noFill_bgPr(self):
        """Return a new `p:bgPr` element with noFill properties."""
        xml = "<p:bgPr %s>\n" "  <a:noFill/>\n" "  <a:effectLst/>\n" "</p:bgPr>" % nsdecls("a", "p")
        bgPr = cast(CT_BackgroundProperties, parse_xml(xml))
        self._insert_bgPr(bgPr)
        return bgPr


class CT_BackgroundProperties(BaseOxmlElement):
    """`p:bgPr` element."""

    _tag_seq = (
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:blipFill",
        "a:pattFill",
        "a:grpFill",
        "a:effectLst",
        "a:effectDag",
        "a:extLst",
    )
    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ),
        successors=_tag_seq[6:],
    )
    del _tag_seq

    def _new_gradFill(self):
        """Override default to add default gradient subtree."""
        return CT_GradientFillProperties.new_gradFill()


class CT_CommonSlideData(BaseOxmlElement):
    """`p:cSld` element."""

    _remove_bg: Callable[[], None]
    get_or_add_bg: Callable[[], CT_Background]

    _tag_seq = ("p:bg", "p:spTree", "p:custDataLst", "p:controls", "p:extLst")
    bg: CT_Background | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bg", successors=_tag_seq[1:]
    )
    spTree: CT_GroupShape = OneAndOnlyOne("p:spTree")  # pyright: ignore[reportAssignmentType]
    del _tag_seq
    name: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString, default=""
    )

    def get_or_add_bgPr(self) -> CT_BackgroundProperties:
        """Return `p:bg/p:bgPr` grandchild.

        If no such grandchild is present, any existing `p:bg` child is first removed and a new
        default `p:bg` with noFill settings is added.
        """
        bg = self.bg
        if bg is None or bg.bgPr is None:
            bg = self._change_to_noFill_bg()
        return cast(CT_BackgroundProperties, bg.bgPr)

    def _change_to_noFill_bg(self) -> CT_Background:
        """Establish a `p:bg` child with no-fill settings.

        Any existing `p:bg` child is first removed.
        """
        self._remove_bg()
        bg = self.get_or_add_bg()
        bg.add_noFill_bgPr()
        return bg


class CT_NotesMaster(_BaseSlideElement):
    """`p:notesMaster` element, root of a notes master part."""

    _tag_seq = ("p:cSld", "p:clrMap", "p:hf", "p:notesStyle", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @classmethod
    def new_default(cls) -> CT_NotesMaster:
        """Return a new `p:notesMaster` element based on the built-in default template."""
        return cast(CT_NotesMaster, parse_from_template("notesMaster"))


class CT_NotesSlide(_BaseSlideElement):
    """`p:notes` element, root of a notes slide part."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @classmethod
    def new(cls) -> CT_NotesSlide:
        """Return a new ``<p:notes>`` element based on the default template.

        Note that the template does not include placeholders, which must be subsequently cloned
        from the notes master.
        """
        return cast(CT_NotesSlide, parse_from_template("notes"))


class CT_Slide(_BaseSlideElement):
    """`p:sld` element, root element of a slide part (XML document)."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    clrMapOvr = ZeroOrOne("p:clrMapOvr", successors=_tag_seq[2:])
    timing = ZeroOrOne("p:timing", successors=_tag_seq[4:])
    del _tag_seq

    def add_animation(self, shape_id: int, animation_type_xml: str, duration: int) -> None:
        """Add an entrance animation for the shape identified by `shape_id`.

        `animation_type_xml` is a string key from `_ANIMATION_FILTER_MAP`,
        `duration` is in milliseconds.
        """
        existing_timing = self.timing
        if existing_timing is None:
            # No timing element yet — create a fresh one
            timing_xml = _build_animation_timing_xml(shape_id, animation_type_xml, duration)
            timing_el = parse_xml(timing_xml)
            self._insert_timing(timing_el)
        else:
            # Add to existing mainSeq childTnLst
            main_seq_childTnLst = self._get_or_add_main_seq_childTnLst()
            next_id = self._next_cTn_id
            par_xml = _build_animation_click_par_xml(
                shape_id, animation_type_xml, duration, next_id
            )
            par_el = parse_xml(par_xml)
            main_seq_childTnLst.append(par_el)

    @property
    def _next_cTn_id(self) -> int:
        """Return the next available unique cTn ID across all timing nodes."""
        from pptx.oxml.ns import qn

        cTn_els = self.findall(".//" + qn("p:cTn"))
        if not cTn_els:
            return 1
        ids = [int(el.get("id", "0")) for el in cTn_els]
        return max(ids) + 1

    def _get_or_add_main_seq_childTnLst(self):
        """Return the `p:childTnLst` inside `p:seq[@nodeType='mainSeq']`.

        Creates the full timing scaffold if it doesn't exist.
        """
        from pptx.oxml.ns import qn

        # Look for existing mainSeq childTnLst
        seq_els = self.findall(".//" + qn("p:seq"))
        for seq_el in seq_els:
            cTn = seq_el.find(qn("p:cTn"))
            if cTn is not None and cTn.get("nodeType") == "mainSeq":
                childTnLst = cTn.find(qn("p:childTnLst"))
                if childTnLst is not None:
                    return childTnLst
        # If no mainSeq exists, create fresh timing scaffold
        if self.timing is not None:
            self.remove(self.timing)
        timing_xml = (
            "<p:timing %s>\n"
            "  <p:tnLst>\n"
            "    <p:par>\n"
            '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">\n'
            "        <p:childTnLst>\n"
            '          <p:seq concurrent="1" nextAc="seek">\n'
            '            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">\n'
            "              <p:childTnLst/>\n"
            "            </p:cTn>\n"
            "            <p:prevCondLst>\n"
            '              <p:cond evt="onPrev" delay="0">\n'
            "                <p:tgtEl>\n"
            "                  <p:sldTgt/>\n"
            "                </p:tgtEl>\n"
            "              </p:cond>\n"
            "            </p:prevCondLst>\n"
            "            <p:nextCondLst>\n"
            '              <p:cond evt="onNext" delay="0">\n'
            "                <p:tgtEl>\n"
            "                  <p:sldTgt/>\n"
            "                </p:tgtEl>\n"
            "              </p:cond>\n"
            "            </p:nextCondLst>\n"
            "          </p:seq>\n"
            "        </p:childTnLst>\n"
            "      </p:cTn>\n"
            "    </p:par>\n"
            "  </p:tnLst>\n"
            "</p:timing>\n" % nsdecls("p")
        )
        timing_el = parse_xml(timing_xml)
        self._insert_timing(timing_el)
        childTnLst_els = self.findall(
            ".//" + qn("p:seq") + "/" + qn("p:cTn") + "/" + qn("p:childTnLst")
        )
        return childTnLst_els[0]

    @classmethod
    def new(cls) -> CT_Slide:
        """Return new `p:sld` element configured as base slide shape."""
        return cast(CT_Slide, parse_xml(cls._sld_xml()))

    @property
    def bg(self):
        """Return `p:bg` grandchild or None if not present."""
        return self.cSld.bg

    def get_or_add_childTnLst(self):
        """Return parent element for a new `p:video` child element.

        The `p:video` element causes play controls to appear under a video
        shape (pic shape containing video). There can be more than one video
        shape on a slide, which causes the precondition to vary. It needs to
        handle the case when there is no `p:sld/p:timing` element and when
        that element already exists. If the case isn't simple, it just nukes
        what's there and adds a fresh one. This could theoretically remove
        desired existing timing information, but there isn't any evidence
        available to me one way or the other, so I've taken the simple
        approach.
        """
        childTnLst = self._childTnLst
        if childTnLst is None:
            childTnLst = self._add_childTnLst()
        return childTnLst

    def _add_childTnLst(self):
        """Add `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Any existing `p:timing` child element is ruthlessly removed and
        replaced.
        """
        self.remove(self.get_or_add_timing())
        timing = parse_xml(self._childTnLst_timing_xml())
        self._insert_timing(timing)
        return timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst")[0]

    @property
    def _childTnLst(self):
        """Return `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Return None if that element is not present.
        """
        childTnLsts = self.xpath("./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst")
        if not childTnLsts:
            return None
        return childTnLsts[0]

    @staticmethod
    def _childTnLst_timing_xml():
        return (
            "<p:timing %s>\n"
            "  <p:tnLst>\n"
            "    <p:par>\n"
            '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="'
            'tmRoot">\n'
            "        <p:childTnLst/>\n"
            "      </p:cTn>\n"
            "    </p:par>\n"
            "  </p:tnLst>\n"
            "</p:timing>" % nsdecls("p")
        )

    @staticmethod
    def _sld_xml():
        return (
            "<p:sld %s>\n"
            "  <p:cSld>\n"
            "    <p:spTree>\n"
            "      <p:nvGrpSpPr>\n"
            '        <p:cNvPr id="1" name=""/>\n'
            "        <p:cNvGrpSpPr/>\n"
            "        <p:nvPr/>\n"
            "      </p:nvGrpSpPr>\n"
            "      <p:grpSpPr/>\n"
            "    </p:spTree>\n"
            "  </p:cSld>\n"
            "  <p:clrMapOvr>\n"
            "    <a:masterClrMapping/>\n"
            "  </p:clrMapOvr>\n"
            "</p:sld>" % nsdecls("a", "p", "r")
        )


class CT_SlideLayout(_BaseSlideElement):
    """`p:sldLayout` element, root of a slide layout part."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq


class CT_SlideLayoutIdList(BaseOxmlElement):
    """`p:sldLayoutIdLst` element, child of `p:sldMaster`.

    Contains references to the slide layouts that inherit from the slide master.
    """

    sldLayoutId_lst: list[CT_SlideLayoutIdListEntry]

    sldLayoutId = ZeroOrMore("p:sldLayoutId")


class CT_SlideLayoutIdListEntry(BaseOxmlElement):
    """`p:sldLayoutId` element, child of `p:sldLayoutIdLst`.

    Contains a reference to a slide layout.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideMaster(_BaseSlideElement):
    """`p:sldMaster` element, root of a slide master part."""

    get_or_add_sldLayoutIdLst: Callable[[], CT_SlideLayoutIdList]

    _tag_seq = (
        "p:cSld",
        "p:clrMap",
        "p:sldLayoutIdLst",
        "p:transition",
        "p:timing",
        "p:hf",
        "p:txStyles",
        "p:extLst",
    )
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    sldLayoutIdLst: CT_SlideLayoutIdList = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldLayoutIdLst", successors=_tag_seq[3:]
    )
    del _tag_seq


class CT_SlideTiming(BaseOxmlElement):
    """`p:timing` element, specifying animations and timed behaviors."""

    _tag_seq = ("p:tnLst", "p:bldLst", "p:extLst")
    tnLst = ZeroOrOne("p:tnLst", successors=_tag_seq[1:])
    del _tag_seq


class CT_TimeNodeList(BaseOxmlElement):
    """`p:tnLst` or `p:childTnList` element."""

    def add_video(self, shape_id):
        """Add a new `p:video` child element for movie having *shape_id*."""
        video_xml = (
            "<p:video %s>\n"
            '  <p:cMediaNode vol="80000">\n'
            '    <p:cTn id="%d" fill="hold" display="0">\n'
            "      <p:stCondLst>\n"
            '        <p:cond delay="indefinite"/>\n'
            "      </p:stCondLst>\n"
            "    </p:cTn>\n"
            "    <p:tgtEl>\n"
            '      <p:spTgt spid="%d"/>\n'
            "    </p:tgtEl>\n"
            "  </p:cMediaNode>\n"
            "</p:video>\n" % (nsdecls("p"), self._next_cTn_id, shape_id)
        )
        video = parse_xml(video_xml)
        self.append(video)

    @property
    def _next_cTn_id(self):
        """Return the next available unique ID (int) for p:cTn element."""
        cTn_id_strs = self.xpath("/p:sld/p:timing//p:cTn/@id")
        ids = [int(id_str) for id_str in cTn_id_strs]
        return max(ids) + 1


class CT_TLMediaNodeVideo(BaseOxmlElement):
    """`p:video` element, specifying video media details."""

    _tag_seq = ("p:cMediaNode",)
    cMediaNode = OneAndOnlyOne("p:cMediaNode")
    del _tag_seq


# -- Animation filter mappings for OOXML `p:animEffect` element ---

_ANIMATION_FILTER_MAP: dict[str, dict[str, str]] = {
    "appear": {"filter": "", "transition": ""},
    "fade": {"filter": "fade", "transition": "in"},
    "fly_from_bottom": {"filter": "wipe(up)", "transition": "in"},
    "fly_from_left": {"filter": "wipe(right)", "transition": "in"},
    "fly_from_right": {"filter": "wipe(left)", "transition": "in"},
    "fly_from_top": {"filter": "wipe(down)", "transition": "in"},
    "wipe_from_bottom": {"filter": "wipe(up)", "transition": "in"},
    "wipe_from_left": {"filter": "wipe(right)", "transition": "in"},
    "wipe_from_right": {"filter": "wipe(left)", "transition": "in"},
    "wipe_from_top": {"filter": "wipe(down)", "transition": "in"},
    "split_h_out": {"filter": "barn(outHorizontal)", "transition": "in"},
    "split_v_out": {"filter": "barn(outVertical)", "transition": "in"},
    "wheel_1": {"filter": "wheel(1)", "transition": "in"},
    "dissolve": {"filter": "dissolve", "transition": "in"},
}


def _build_animation_timing_xml(shape_id: int, animation_type_xml: str, duration: int) -> str:
    """Build the full `p:timing` XML for an entrance animation on a shape.

    `shape_id` identifies the target shape, `animation_type_xml` is a key
    in `_ANIMATION_FILTER_MAP`, and `duration` is in milliseconds.
    """
    effect_info = _ANIMATION_FILTER_MAP[animation_type_xml]
    filter_str = effect_info["filter"]

    # Build the inner animation node(s)
    if not filter_str:
        # APPEAR animation - just set visibility, no p:animEffect
        effect_nodes_xml = ""
    else:
        effect_nodes_xml = (
            '                            <p:animEffect transition="%s" filter="%s">\n'
            "                              <p:cBhvr>\n"
            '                                <p:cTn id="6" dur="%d" fill="hold"/>\n'
            "                                <p:tgtEl>\n"
            '                                  <p:spTgt spid="%d"/>\n'
            "                                </p:tgtEl>\n"
            "                              </p:cBhvr>\n"
            "                            </p:animEffect>\n"
            % (
                effect_info["transition"],
                filter_str,
                duration,
                shape_id,
            )
        )

    timing_xml = (
        "<p:timing %s>\n"
        "  <p:tnLst>\n"
        "    <p:par>\n"
        '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">\n'
        "        <p:childTnLst>\n"
        '          <p:seq concurrent="1" nextAc="seek">\n'
        '            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">\n'
        "              <p:childTnLst>\n"
        "                <p:par>\n"
        '                  <p:cTn id="3" fill="hold">\n'
        "                    <p:stCondLst>\n"
        '                      <p:cond delay="indefinite"/>\n'
        "                    </p:stCondLst>\n"
        "                    <p:childTnLst>\n"
        "                      <p:par>\n"
        '                        <p:cTn id="4" fill="hold">\n'
        "                          <p:stCondLst>\n"
        '                            <p:cond delay="0"/>\n'
        "                          </p:stCondLst>\n"
        "                          <p:childTnLst>\n"
        "                            <p:set>\n"
        "                              <p:cBhvr>\n"
        '                                <p:cTn id="5" dur="1" fill="hold">\n'
        "                                  <p:stCondLst>\n"
        '                                    <p:cond delay="0"/>\n'
        "                                  </p:stCondLst>\n"
        "                                </p:cTn>\n"
        "                                <p:tgtEl>\n"
        '                                  <p:spTgt spid="%d"/>\n'
        "                                </p:tgtEl>\n"
        "                                <p:attrNameLst>\n"
        "                                  <p:attrName>style.visibility</p:attrName>\n"
        "                                </p:attrNameLst>\n"
        "                              </p:cBhvr>\n"
        "                              <p:to>\n"
        '                                <p:strVal val="visible"/>\n'
        "                              </p:to>\n"
        "                            </p:set>\n"
        "%s"
        "                          </p:childTnLst>\n"
        "                        </p:cTn>\n"
        "                      </p:par>\n"
        "                    </p:childTnLst>\n"
        "                  </p:cTn>\n"
        "                </p:par>\n"
        "              </p:childTnLst>\n"
        "            </p:cTn>\n"
        "            <p:prevCondLst>\n"
        '              <p:cond evt="onPrev" delay="0">\n'
        "                <p:tgtEl>\n"
        "                  <p:sldTgt/>\n"
        "                </p:tgtEl>\n"
        "              </p:cond>\n"
        "            </p:prevCondLst>\n"
        "            <p:nextCondLst>\n"
        '              <p:cond evt="onNext" delay="0">\n'
        "                <p:tgtEl>\n"
        "                  <p:sldTgt/>\n"
        "                </p:tgtEl>\n"
        "              </p:cond>\n"
        "            </p:nextCondLst>\n"
        "          </p:seq>\n"
        "        </p:childTnLst>\n"
        "      </p:cTn>\n"
        "    </p:par>\n"
        "  </p:tnLst>\n"
        "</p:timing>\n"
        % (nsdecls("p"), shape_id, effect_nodes_xml)
    )

    return timing_xml


def _build_animation_click_par_xml(
    shape_id: int, animation_type_xml: str, duration: int, start_cTn_id: int
) -> str:
    """Build a `p:par` XML fragment for one click-animation sequence item.

    Used when adding an animation to a slide that already has animations.
    `start_cTn_id` should be the next available cTn id.
    """
    effect_info = _ANIMATION_FILTER_MAP[animation_type_xml]
    filter_str = effect_info["filter"]

    # Allocate cTn IDs: outer par, inner par, set's cTn, (optionally) animEffect's cTn
    id_outer_par = start_cTn_id
    id_inner_par = start_cTn_id + 1
    id_set_ctn = start_cTn_id + 2

    if not filter_str:
        effect_nodes_xml = ""
    else:
        id_anim_ctn = start_cTn_id + 3
        effect_nodes_xml = (
            '                            <p:animEffect transition="%s" filter="%s">\n'
            "                              <p:cBhvr>\n"
            '                                <p:cTn id="%d" dur="%d" fill="hold"/>\n'
            "                                <p:tgtEl>\n"
            '                                  <p:spTgt spid="%d"/>\n'
            "                                </p:tgtEl>\n"
            "                              </p:cBhvr>\n"
            "                            </p:animEffect>\n"
            % (
                effect_info["transition"],
                filter_str,
                id_anim_ctn,
                duration,
                shape_id,
            )
        )

    par_xml = (
        "<p:par %s>\n"
        '  <p:cTn id="%d" fill="hold">\n'
        "    <p:stCondLst>\n"
        '      <p:cond delay="indefinite"/>\n'
        "    </p:stCondLst>\n"
        "    <p:childTnLst>\n"
        "      <p:par>\n"
        '        <p:cTn id="%d" fill="hold">\n'
        "          <p:stCondLst>\n"
        '            <p:cond delay="0"/>\n'
        "          </p:stCondLst>\n"
        "          <p:childTnLst>\n"
        "            <p:set>\n"
        "              <p:cBhvr>\n"
        '                <p:cTn id="%d" dur="1" fill="hold">\n'
        "                  <p:stCondLst>\n"
        '                    <p:cond delay="0"/>\n'
        "                  </p:stCondLst>\n"
        "                </p:cTn>\n"
        "                <p:tgtEl>\n"
        '                  <p:spTgt spid="%d"/>\n'
        "                </p:tgtEl>\n"
        "                <p:attrNameLst>\n"
        "                  <p:attrName>style.visibility</p:attrName>\n"
        "                </p:attrNameLst>\n"
        "              </p:cBhvr>\n"
        "              <p:to>\n"
        '                <p:strVal val="visible"/>\n'
        "              </p:to>\n"
        "            </p:set>\n"
        "%s"
        "          </p:childTnLst>\n"
        "        </p:cTn>\n"
        "      </p:par>\n"
        "    </p:childTnLst>\n"
        "  </p:cTn>\n"
        "</p:par>\n"
        % (nsdecls("p"), id_outer_par, id_inner_par, id_set_ctn, shape_id, effect_nodes_xml)
    )

    return par_xml
