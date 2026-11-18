using System;
using System.Collections.Generic;

namespace CalcRapido
{
    // Detección de límites de compatibilidad al abrir un libro. No bloquea la
    // carga: solo recoge funciones desconocidas para que la interfaz avise.
    public static class CompatibilityScanner
    {
        public static void ScanFormulas(Workbook wb)
        {
            if (wb == null) return;
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Sheet sheet in wb.Sheets)
            {
                foreach (Cell cell in sheet.Cells.Values)
                {
                    if (!cell.IsFormula || string.IsNullOrEmpty(cell.Formula)) continue;
                    try
                    {
                        List<Token> tokens = new FormulaLexer(cell.Formula).Tokenize();
                        for (int i = 0; i + 1 < tokens.Count; i++)
                        {
                            if (tokens[i].Kind != TokenKind.Ident ||
                                tokens[i + 1].Kind != TokenKind.LParen) continue;
                            string name = tokens[i].Text;
                            if (string.IsNullOrEmpty(name)) continue;
                            if (Builtins.Canonical(name) == null) seen.Add(name);
                        }
                    }
                    catch (FormulaParseException)
                    {
                        // Una fórmula inválida se cuenta como no compatible,
                        // pero no impide abrir el libro.
                        seen.Add("fórmula no reconocida");
                    }
                    catch
                    {
                        // Nunca permitir que el diagnóstico impida la carga.
                    }
                }
            }
            foreach (string name in seen) wb.AddUnsupportedFunction(name);
            if (seen.Count > 0)
                wb.AddLoadWarning("Algunas fórmulas usan funciones que esta versión no calcula. " +
                                  "Se conserva el valor guardado cuando existe; si editas la fórmula puede aparecer un error.");
        }
    }
}
