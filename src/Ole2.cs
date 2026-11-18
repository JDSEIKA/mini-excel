using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CalcRapido
{
    // Lector mínimo de contenedores OLE2/Compound File Binary (los .xls
    // antiguos guardan el flujo "Workbook" dentro de este contenedor).
    internal static class Ole2
    {
        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FreeSector = 0xFFFFFFFF;
        private const long MaxFileBytes = 120L * 1024 * 1024;
        private const long MaxOutputBytes = 256L * 1024 * 1024;

        public static byte[] ReadWorkbookStream(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (fi.Length > MaxFileBytes)
                throw new FileFormatException("El archivo .xls es demasiado grande.");
            byte[] data = File.ReadAllBytes(path);
            if (data.Length < 512 || data[0] != 0xD0 || data[1] != 0xCF ||
                data[2] != 0x11 || data[3] != 0xE0)
                throw new FileFormatException("El archivo no es un libro .xls válido.");
            int sectorShift = U16(data, 30);
            int miniShift = U16(data, 32);
            if (sectorShift != 9 && sectorShift != 12)
                throw new FileFormatException("Contenedor .xls no válido.");
            if (miniShift != 6)
                throw new FileFormatException("Contenedor .xls no válido.");
            int sectorSize = 1 << sectorShift;
            int miniSectorSize = 1 << miniShift;
            uint numFat = U32(data, 44);
            uint dirStart = U32(data, 48);
            uint miniCutoff = U32(data, 56);
            uint firstMiniFat = U32(data, 60);
            uint numMiniFat = U32(data, 64);
            uint firstDifat = U32(data, 68);
            uint numDifat = U32(data, 72);
            if (sectorSize < 512 || sectorSize > 65536) throw new FileFormatException("Contenedor .xls no válido.");

            List<uint> fatSectors = new List<uint>();
            for (int i = 0; i < 109; i++)
            {
                uint s = U32(data, 76 + i * 4);
                if (s < 0xFFFFFFFA) fatSectors.Add(s);
            }
            uint difat = firstDifat;
            HashSet<uint> seenDifat = new HashSet<uint>();
            uint difatSteps = 0;
            while (difat < 0xFFFFFFFA && difatSteps++ < numDifat && difatSteps <= 10000)
            {
                if (!seenDifat.Add(difat))
                    throw new FileFormatException("Contenedor .xls con cadena DIFAT circular.");
                byte[] sector = Sector(data, difat, sectorSize);
                int per = sectorSize / 4 - 1;
                for (int i = 0; i < per; i++)
                {
                    uint s = U32(sector, i * 4);
                    if (s < 0xFFFFFFFA)
                    {
                        fatSectors.Add(s);
                        if (fatSectors.Count > 100000)
                            throw new FileFormatException("Contenedor .xls demasiado grande.");
                    }
                }
                difat = U32(sector, per * 4);
            }

            List<uint> fat = new List<uint>();
            foreach (uint fs in fatSectors)
            {
                byte[] sector = Sector(data, fs, sectorSize);
                int count = sectorSize / 4;
                for (int i = 0; i < count; i++) fat.Add(U32(sector, i * 4));
            }
            if (fat.Count == 0) throw new FileFormatException("Contenedor .xls sin tabla FAT.");

            byte[] dirBytes = ReadChain(data, fat, dirStart, -1, sectorSize);
            List<DirEntry> entries = ParseDirectory(dirBytes);
            DirEntry workbook = null;
            DirEntry root = null;
            foreach (DirEntry e in entries)
            {
                if (e.Type == 5) root = e;
                if (e.Type == 2 && (string.Equals(e.Name, "Workbook", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(e.Name, "Book", StringComparison.OrdinalIgnoreCase)))
                {
                    if (workbook == null || string.Equals(e.Name, "Workbook", StringComparison.OrdinalIgnoreCase))
                        workbook = e;
                }
            }
            if (workbook == null) throw new FileFormatException("El archivo .xls no contiene el flujo Workbook.");
            if (workbook.Size == 0 || workbook.Size > MaxOutputBytes)
                throw new FileFormatException("El flujo Workbook del .xls no es válido o es demasiado grande.");

            byte[] miniFat = new byte[0];
            if (numMiniFat > 0 && firstMiniFat < 0xFFFFFFFA)
            {
                if (numMiniFat > 100000 || (long)numMiniFat * sectorSize > MaxOutputBytes)
                    throw new FileFormatException("Contenedor .xls demasiado grande.");
                int miniFatBytes = checked((int)numMiniFat * sectorSize);
                miniFat = ReadChain(data, fat, firstMiniFat, miniFatBytes, sectorSize);
            }
            byte[] miniStream = new byte[0];
            if (root != null && root.Size > 0)
            {
                if (root.Size > MaxOutputBytes)
                    throw new FileFormatException("Contenedor .xls demasiado grande.");
                miniStream = ReadChain(data, fat, root.Start, (int)root.Size, sectorSize);
            }

            if (workbook.Size < miniCutoff && miniStream.Length > 0)
            {
                List<uint> mfat = new List<uint>();
                for (int i = 0; i + 4 <= miniFat.Length; i += 4) mfat.Add(U32(miniFat, i));
                return ReadMiniChain(miniStream, mfat, workbook.Start, (int)Math.Min(workbook.Size, int.MaxValue), miniSectorSize);
            }
            return ReadChain(data, fat, workbook.Start, (int)Math.Min(workbook.Size, int.MaxValue), sectorSize);
        }

        private sealed class DirEntry
        {
            public string Name;
            public int Type;
            public uint Start;
            public uint Size;
        }

        private static List<DirEntry> ParseDirectory(byte[] bytes)
        {
            List<DirEntry> list = new List<DirEntry>();
            for (int pos = 0; pos + 128 <= bytes.Length; pos += 128)
            {
                int nameLen = U16(bytes, pos + 64);
                if (nameLen < 2 || nameLen > 64) { list.Add(new DirEntry { Name = "", Type = 0 }); continue; }
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < nameLen - 2; i += 2)
                    sb.Append((char)U16(bytes, pos + i));
                list.Add(new DirEntry
                {
                    Name = sb.ToString(),
                    Type = bytes[pos + 66],
                    Start = U32(bytes, pos + 116),
                    Size = U32(bytes, pos + 120)
                });
            }
            return list;
        }

        private static byte[] ReadChain(byte[] data, List<uint> fat, uint start, int size, int sectorSize)
        {
            MemoryStream ms = new MemoryStream();
            HashSet<uint> seen = new HashSet<uint>();
            uint sector = start;
            while (sector < 0xFFFFFFFA && sector < fat.Count)
            {
                if (!seen.Add(sector))
                    throw new FileFormatException("Contenedor .xls con cadena de sectores circular.");
                int off = checked((int)((sector + 1) * (uint)sectorSize));
                if (off < 0 || off >= data.Length) break;
                int n = Math.Min(sectorSize, data.Length - off);
                if (size >= 0) n = Math.Min(n, size - (int)ms.Length);
                if (n <= 0) break;
                if ((long)ms.Length + n > MaxOutputBytes)
                    throw new FileFormatException("Contenedor .xls demasiado grande.");
                ms.Write(data, off, n);
                if (size >= 0 && ms.Length >= size) break;
                sector = fat[(int)sector];
            }
            return ms.ToArray();
        }

        private static byte[] ReadMiniChain(byte[] miniStream, List<uint> miniFat, uint start, int size, int miniSize)
        {
            MemoryStream ms = new MemoryStream();
            HashSet<uint> seen = new HashSet<uint>();
            uint sector = start;
            while (sector < 0xFFFFFFFA && sector < miniFat.Count)
            {
                if (!seen.Add(sector))
                    throw new FileFormatException("Contenedor .xls con cadena de mini-sectores circular.");
                int off = checked((int)(sector * (uint)miniSize));
                if (off < 0 || off >= miniStream.Length) break;
                int n = Math.Min(miniSize, miniStream.Length - off);
                if (size >= 0) n = Math.Min(n, size - (int)ms.Length);
                if (n <= 0) break;
                if ((long)ms.Length + n > MaxOutputBytes)
                    throw new FileFormatException("Contenedor .xls demasiado grande.");
                ms.Write(miniStream, off, n);
                if (size >= 0 && ms.Length >= size) break;
                sector = miniFat[(int)sector];
            }
            return ms.ToArray();
        }

        private static byte[] Sector(byte[] data, uint sector, int size)
        {
            int off = checked((int)((sector + 1) * (uint)size));
            if (off < 0 || off > data.Length - size) throw new FileFormatException("Contenedor .xls truncado.");
            byte[] buf = new byte[size];
            Buffer.BlockCopy(data, off, buf, 0, size);
            return buf;
        }

        private static ushort U16(byte[] b, int p)
        {
            if (p < 0 || p + 2 > b.Length) return 0;
            return (ushort)(b[p] | (b[p + 1] << 8));
        }

        private static uint U32(byte[] b, int p)
        {
            if (p < 0 || p + 4 > b.Length) return 0;
            return (uint)(b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24));
        }
    }
}
