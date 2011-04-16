using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pictures = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Foxby.Core.MetaObjects
{
    public class Image
    {
        private const string SchemaUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private const string PictureName = "Image inserted via Foxby";
        private const string BlipUriGuid = "{28A0092B-C50C-407E-A947-70E740481C1C}";

        private readonly string imagePartId;

        public Image(string imagePartId)
        {
            this.imagePartId = imagePartId;
        }

        public OpenXmlElement ToOpenXml()
        {
            var drawing =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent {Cx = 990000L, Cy = 792000L},
                        new DW.EffectExtent
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                        new DW.DocProperties
                            {
                                Id = (UInt32Value) 1U,
                                Name = "Picture"
                            },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks {NoChangeAspect = true}),
                        new A.Graphic(
                            new A.GraphicData(
                                new Pictures.Picture(
                                    new Pictures.NonVisualPictureProperties(
                                        new Pictures.NonVisualDrawingProperties
                                            {
                                                Id = (UInt32Value) 0U,
                                                Name = PictureName
                                            },
                                        new Pictures.NonVisualPictureDrawingProperties()),
                                    new Pictures.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension
                                                    {
                                                        Uri = BlipUriGuid
                                                    })
                                            )
                                            {
                                                Embed = imagePartId,
                                                CompressionState =
                                                    A.BlipCompressionValues.Print
                                            },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new Pictures.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset {X = 0L, Y = 0L},
                                            new A.Extents {Cx = 990000L, Cy = 792000L}),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList()
                                            ) {Preset = A.ShapeTypeValues.Rectangle}))
                                ) {Uri = SchemaUri})
                        )
                        {
                            DistanceFromTop = (UInt32Value) 0U,
                            DistanceFromBottom = (UInt32Value) 0U,
                            DistanceFromLeft = (UInt32Value) 0U,
                            DistanceFromRight = (UInt32Value) 0U,
                        });

            return new Run(drawing);
        }
    }
}