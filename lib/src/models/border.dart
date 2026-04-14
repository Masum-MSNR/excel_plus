part of '../../excel_plus.dart';

class Border {
  final BorderStyle? borderStyle;
  final String? borderColorHex;

  Border({BorderStyle? borderStyle, ExcelColor? borderColorHex})
      : borderStyle = borderStyle == BorderStyle.None ? null : borderStyle,
        borderColorHex = borderColorHex != null
            ? _isColorAppropriate(borderColorHex.colorHex)
            : null;

  @override
  String toString() {
    return 'Border(borderStyle: $borderStyle, borderColorHex: $borderColorHex)';
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is Border &&
          other.borderStyle == borderStyle &&
          other.borderColorHex == borderColorHex;

  @override
  int get hashCode => Object.hash(borderStyle, borderColorHex);
}

class _BorderSet {
  final Border leftBorder;
  final Border rightBorder;
  final Border topBorder;
  final Border bottomBorder;
  final Border diagonalBorder;
  final bool diagonalBorderUp;
  final bool diagonalBorderDown;

  _BorderSet({
    required this.leftBorder,
    required this.rightBorder,
    required this.topBorder,
    required this.bottomBorder,
    required this.diagonalBorder,
    required this.diagonalBorderUp,
    required this.diagonalBorderDown,
  });

  _BorderSet copyWith({
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool? diagonalBorderUp,
    bool? diagonalBorderDown,
  }) {
    return _BorderSet(
      leftBorder: leftBorder ?? this.leftBorder,
      rightBorder: rightBorder ?? this.rightBorder,
      topBorder: topBorder ?? this.topBorder,
      bottomBorder: bottomBorder ?? this.bottomBorder,
      diagonalBorder: diagonalBorder ?? this.diagonalBorder,
      diagonalBorderUp: diagonalBorderUp ?? this.diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDown ?? this.diagonalBorderDown,
    );
  }

  @override
  bool operator ==(Object other) =>
      identical(this, other) ||
      other is _BorderSet &&
          other.leftBorder == leftBorder &&
          other.rightBorder == rightBorder &&
          other.topBorder == topBorder &&
          other.bottomBorder == bottomBorder &&
          other.diagonalBorder == diagonalBorder &&
          other.diagonalBorderUp == diagonalBorderUp &&
          other.diagonalBorderDown == diagonalBorderDown;

  @override
  int get hashCode => Object.hash(
        leftBorder,
        rightBorder,
        topBorder,
        bottomBorder,
        diagonalBorder,
        diagonalBorderUp,
        diagonalBorderDown,
      );
}

enum BorderStyle {
  None('none'),
  DashDot('dashDot'),
  DashDotDot('dashDotDot'),
  Dashed('dashed'),
  Dotted('dotted'),
  Double('double'),
  Hair('hair'),
  Medium('medium'),
  MediumDashDot('mediumDashDot'),
  MediumDashDotDot('mediumDashDotDot'),
  MediumDashed('mediumDashed'),
  SlantDashDot('slantDashDot'),
  Thick('thick'),
  Thin('thin');

  final String style;
  const BorderStyle(this.style);
}

BorderStyle? getBorderStyleByName(String name) {
  final lower = 'borderstyle.${name.toLowerCase()}';
  for (final e in BorderStyle.values) {
    if (e.toString().toLowerCase() == lower) return e;
  }
  return null;
}
