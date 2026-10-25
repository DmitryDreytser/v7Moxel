using System;

namespace Moxel
{
    /// <summary>
    /// Реализация IDisposable для модели Moxel (вынесена из основного файла,
    /// чтобы Moxel.cs содержал только логику модели и чтения формата MOXEL).
    /// </summary>
    public partial class Moxel
    {
        private bool _disposed;

        /// <summary>Признак того, что объект уже освобождён.</summary>
        public bool IsDisposed => _disposed;

        /// <summary>
        /// Освобождает ресурсы: кэшированные изображения OLE-объектов и ссылки
        /// на большие коллекции модели.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            // Подавление финализатора не требуется: ~Moxel() пустой.
        }

        void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;

                if (Objects != null)
                {
                    foreach (var obj in Objects)
                    {
                        obj.pObject?.Dispose();
                    }
                }

                FormatsList = null;
                stringTable = null;
                FontList = null;
                Columns = null;
                Rows = null;
                Objects = null;
                Unions = null;
                VerticalSections = null;
                HorisontalSections = null;
                HorisontalPageBreaks = null;
                VerticalPageBreaks = null;
                AreaNames = null;
            }
        }
    }
}
