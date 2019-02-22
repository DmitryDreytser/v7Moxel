using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    [Guid("38E77354-4165-11D6-9AF6-0080AD7A3F21")]
    [InterfaceType(2)]
    [TypeLibType(4112)]
    public interface _DBmp_1c
    {
        [DispId(2)]
        string BmpFile { get; set; }
        [DispId(10)]
        short DstDeltaPointX { get; set; }
        [DispId(11)]
        short DstDeltaPointY { get; set; }
        [DispId(13)]
        short DstHeight { get; set; }
        [DispId(12)]
        short DstWidth { get; set; }
        [DispId(4)]
        short Function { get; set; }
        [DispId(5)]
        short GrMode { get; set; }
        [DispId(3)]
        short GroundClip { get; set; }
        [DispId(15)]
        short GroundGor { get; set; }
        [DispId(14)]
        short GroundVer { get; set; }
        [DispId(1)]
        short NoDraw { get; set; }
        [DispId(-525)]
        int ReadyState { get; }
        [DispId(9)]
        short SrcHeight { get; set; }
        [DispId(6)]
        short SrcPointX { get; set; }
        [DispId(7)]
        short SrcPointY { get; set; }
        [DispId(8)]
        short SrcWidth { get; set; }
    }
}