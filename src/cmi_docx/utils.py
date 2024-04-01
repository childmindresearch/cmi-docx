"""Utility functions for the cmi_docx package."""


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Converts RGB values to a hexadecimal color code.

    Args:
        r: The red component of the RGB color.
        g: The green component of the RGB color.
        b: The blue component of the RGB color.

    Returns:
        The hexadecimal color code representing the RGB color.
    """
    return f"#{r:02x}{g:02x}{b:02x}".upper()
