using System;
using System.IO;

namespace CalcRapido
{
    /// <summary>
    /// Escribe primero en un archivo temporal de la misma carpeta y solo
    /// sustituye el destino cuando la escritura terminó correctamente.
    /// </summary>
    internal static class AtomicIO
    {
        public static void Save(string path, Action<string> writer)
        {
            if (writer == null) throw new ArgumentNullException("writer");
            string full = Path.GetFullPath(path);
            string dir = Path.GetDirectoryName(full);
            if (string.IsNullOrEmpty(dir)) throw new IOException("Ruta de destino no válida.");
            Directory.CreateDirectory(dir);
            string ext = Path.GetExtension(full);
            string temp = Path.Combine(dir,
                "." + Path.GetFileNameWithoutExtension(full) + "." +
                Guid.NewGuid().ToString("N") + ext);
            bool replaced = false;
            try
            {
                writer(temp);
                if (!File.Exists(temp))
                    throw new IOException("El escritor no generó el archivo temporal.");
                if (File.Exists(full))
                {
                    try
                    {
                        File.Replace(temp, full, null, true);
                    }
                    catch (PlatformNotSupportedException ex)
                    {
                        throw new IOException("El sistema de archivos no permite el reemplazo atómico.", ex);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException("No se pudo reemplazar el archivo de forma atómica. " +
                            "El original no fue modificado.", ex);
                    }
                }
                else
                {
                    File.Move(temp, full);
                }
                replaced = true;
            }
            finally
            {
                if (!replaced)
                {
                    try { if (File.Exists(temp)) File.Delete(temp); } catch { }
                }
            }
        }
    }
}
