# textbox constants
WIDTH_BASE = 184666
HEIGHT_BASE = 92333

# fontsize constants
WIDTH_CONSOLAS_18PT = 126638
HEIGHT_CONSOLAS_18PT = 276999

# slide constants
SLIDE_WIDTH = 9144000
SLIDE_HEIGHT = 6858000
CENTER = (SLIDE_WIDTH // 2, SLIDE_HEIGHT // 2)


def nextline(y, is_box=True):
    if is_box:
        return y + HEIGHT_BASE + HEIGHT_CONSOLAS_18PT
    else:
        return y + HEIGHT_CONSOLAS_18PT


def anchor(center_x, center_y, width, height):
    """
    Return coordinates of the upper left corner
    """
    x = center_x - width // 2
    y = center_y - height // 2
    return x, y


def get_width_height(text, width_pad=5):
    """
    width_pad :
        unit character width
    """
    max_len = max(len(l) for l in text.split("\n"))
    n_lines = text.count("\n") + 1
    width = WIDTH_BASE + (max_len + width_pad) * WIDTH_CONSOLAS_18PT
    height = HEIGHT_BASE + n_lines * HEIGHT_CONSOLAS_18PT
    return width, height
