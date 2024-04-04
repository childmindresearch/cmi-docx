"""Test for utility functions."""

import pytest

from cmi_docx import table


@pytest.mark.parametrize(
    ("rgb", "hexadecimal"),
    [
        ((0, 0, 0), "#000000"),
        ((255, 255, 255), "#FFFFFF"),
        ((255, 0, 0), "#FF0000"),
        ((0, 255, 0), "#00FF00"),
        ((0, 0, 255), "#0000FF"),
        ((255, 255, 0), "#FFFF00"),
        ((0, 255, 255), "#00FFFF"),
        ((255, 0, 255), "#FF00FF"),
        ((128, 128, 128), "#808080"),
        ((255, 128, 128), "#FF8080"),
        ((128, 255, 128), "#80FF80"),
        ((128, 128, 255), "#8080FF"),
        ((255, 255, 128), "#FFFF80"),
        ((128, 255, 255), "#80FFFF"),
        ((255, 128, 255), "#FF80FF"),
    ],
)
def test_rgb_to_hex(
    rgb: tuple[int, int, int],
    hexadecimal: str,
) -> None:
    """Tests converting RGB to hex."""
    assert table.rgb_to_hex(*rgb) == hexadecimal
