import 'package:flutter_test/flutter_test.dart';

import 'package:editor_pdf/main.dart';

void main() {
  testWidgets('La app carga con el título de conversión', (tester) async {
    await tester.pumpWidget(const EditorPdfApp());
    expect(find.text('Convertir PDF a Word'), findsOneWidget);
    expect(find.text('Elegir archivo PDF'), findsOneWidget);
  });
}
