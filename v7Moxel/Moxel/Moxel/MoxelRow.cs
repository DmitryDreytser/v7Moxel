using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using static Moxel.Moxel;
using System.Collections;

namespace Moxel
{
    [Serializable]
    public class MoxelRow : IDictionary<int, DataCell>
    {
        Moxel Parent = null;
        public CSheetFormat FormatCell;
        public Dictionary<int, DataCell> values;

        public int Height
        {
            get
            {
                if (FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                    if (FormatCell.wHeight == 0)
                        return 0;
                    else
                        return FormatCell.wHeight;
                else
                {
                    foreach (var cell in values.Values)
                    {
                        if (cell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.RowHeight))
                        {
                            return cell.FormatCell.wHeight;
                        }

                        if (string.IsNullOrEmpty(cell.Text))
                            return 0;

                        if (string.IsNullOrEmpty(cell.Text) && cell.FormatCell.dwFlags.HasFlag(MoxelCellFlags.FontSize))
                        {
                            return (int)Math.Floor(cell.FormatCell.wFontSize * (-1.3));
                        }
                    }
                    return 0;
                }
                    
            }
            set
            {
                FormatCell.wHeight = (short)value;
                if (value != 0)
                    FormatCell.dwFlags |= MoxelCellFlags.RowHeight;
                else
                    FormatCell.dwFlags ^= MoxelCellFlags.RowHeight;
            }
        }

        public ICollection<int> Keys
        {
            get
            {
                return values.Keys;
            }
        }

        public ICollection<DataCell> Values
        {
            get
            {
                return values.Values;
            }
        }

        public int Count
        {
            get
            {
                return Parent.nAllColumnCount;
            }
        }

        public bool IsReadOnly
        {
            get
            {
                return false;
            }
        }

        public DataCell this[int key]
        {
            get
            {
                if (values.ContainsKey(key))
                    return values[key];
                else
                    return new DataCell(FormatCell, Parent);
            }

            set
            {
                values[key] = value;
            }
        }

        public MoxelRow()
        { }

        public MoxelRow(BinaryReader br, Moxel parent)
        {
            Parent = parent;
            FormatCell = br.Read<DataCell>(parent);
            values = br.ReadDictionary<DataCell>(parent);
        }

        public bool ContainsKey(int key)
        {
            return values.ContainsKey(key);
        }

        public void Add(int key, DataCell value)
        {
            values.Add(key, value);
        }

        public bool Remove(int key)
        {
            return values.Remove(key);
        }

        public bool TryGetValue(int key, out DataCell value)
        {
            return values.TryGetValue(key, out value);
        }

        public IEnumerator GetEnumerator()
        {
            return Values.GetEnumerator();
        }

        public void Add(KeyValuePair<int, DataCell> item)
        {
            values.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            values.Clear();
        }

        public bool Contains(KeyValuePair<int, DataCell> item)
        {
            return values.Contains(item);
        }

        public void CopyTo(KeyValuePair<int, DataCell>[] array, int arrayIndex)
        {
            ((ICollection)values).CopyTo(array, arrayIndex);
        }

        public bool Remove(KeyValuePair<int, DataCell> item)
        {
            return values.Remove(item.Key);
        }

        IEnumerator<KeyValuePair<int, DataCell>> IEnumerable<KeyValuePair<int, DataCell>>.GetEnumerator()
        {
            return values.GetEnumerator();
        }
    }

}
