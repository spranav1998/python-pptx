"""Enumerations used by animation objects."""

from __future__ import annotations

from pptx.enum.base import BaseXmlEnum


class PP_ANIMATION_TYPE(BaseXmlEnum):
    """Specifies the type of animation effect to apply to a shape.

    Example::

        from pptx.enum.animation import PP_ANIMATION_TYPE

        slide.add_animation(shape, PP_ANIMATION_TYPE.FADE)
    """

    APPEAR = (
        1,
        "appear",
        "Shape appears instantly on click.",
    )
    """Shape appears instantly on click."""

    FADE = (
        2,
        "fade",
        "Shape fades in gradually.",
    )
    """Shape fades in gradually."""

    FLY_IN_FROM_BOTTOM = (
        3,
        "fly_from_bottom",
        "Shape flies in from the bottom of the slide.",
    )
    """Shape flies in from the bottom of the slide."""

    FLY_IN_FROM_LEFT = (
        4,
        "fly_from_left",
        "Shape flies in from the left of the slide.",
    )
    """Shape flies in from the left of the slide."""

    FLY_IN_FROM_RIGHT = (
        5,
        "fly_from_right",
        "Shape flies in from the right of the slide.",
    )
    """Shape flies in from the right of the slide."""

    FLY_IN_FROM_TOP = (
        6,
        "fly_from_top",
        "Shape flies in from the top of the slide.",
    )
    """Shape flies in from the top of the slide."""

    WIPE_FROM_BOTTOM = (
        7,
        "wipe_from_bottom",
        "Shape is revealed with a wipe from the bottom.",
    )
    """Shape is revealed with a wipe from the bottom."""

    WIPE_FROM_LEFT = (
        8,
        "wipe_from_left",
        "Shape is revealed with a wipe from the left.",
    )
    """Shape is revealed with a wipe from the left."""

    WIPE_FROM_RIGHT = (
        9,
        "wipe_from_right",
        "Shape is revealed with a wipe from the right.",
    )
    """Shape is revealed with a wipe from the right."""

    WIPE_FROM_TOP = (
        10,
        "wipe_from_top",
        "Shape is revealed with a wipe from the top.",
    )
    """Shape is revealed with a wipe from the top."""

    SPLIT_HORIZONTAL_OUT = (
        11,
        "split_h_out",
        "Shape is revealed with a horizontal split outward.",
    )
    """Shape is revealed with a horizontal split outward."""

    SPLIT_VERTICAL_OUT = (
        12,
        "split_v_out",
        "Shape is revealed with a vertical split outward.",
    )
    """Shape is revealed with a vertical split outward."""

    WHEEL_1_SPOKE = (
        13,
        "wheel_1",
        "Shape is revealed with a 1-spoke wheel animation.",
    )
    """Shape is revealed with a 1-spoke wheel animation."""

    DISSOLVE = (
        14,
        "dissolve",
        "Shape dissolves in gradually.",
    )
    """Shape dissolves in gradually."""


PP_ANIMATION = PP_ANIMATION_TYPE
