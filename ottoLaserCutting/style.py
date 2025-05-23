from openpyxl.styles import DEFAULT_FONT, Font, NamedStyle, GradientFill, PatternFill, Font, Alignment
from openpyxl.styles.borders import Border, Side

alCenter     = Alignment(horizontal = "center", vertical = "center")
alCenterWrap = Alignment(horizontal = "center", vertical = "center", wrapText = True)
font = {
    # "red": Font(color="ff0000"),
    # "green": Font(color="00b050"),
    # "white":  Font(color="00b050"),
    # "blueUnderline": Font(color="0000ff", underline="single")
    "strikethrough": Font(strike=True),
    "blueStrikethrough": Font(color="0000ff", strike=True),
    "blue": Font(color="0000ff"),
    "orangeBold": Font(color="FFC000", bold=True),
    "greenBold": Font(color="00B050", bold=True)
}

style = {
    "strikethroughtHyperlink": NamedStyle(name="strikethroughHyperlink"),
    "centerHyperlink": NamedStyle(name="centerHyperlink"),
    "centerStrikethrough": NamedStyle(name="centerStrikethrough"),
    "input": NamedStyle(name="input"),
}

borderMedium = Border(top=Side(style="medium"))

style["strikethroughtHyperlink"].font = font["blueStrikethrough"]
style["strikethroughtHyperlink"].alignment = alCenter
style["centerHyperlink"].font = font["blue"]
style["centerHyperlink"].alignment = alCenter
style["centerStrikethrough"].font = font["strikethrough"]
style["centerStrikethrough"].alignment = alCenter
style["input"].font = font["greenBold"]

