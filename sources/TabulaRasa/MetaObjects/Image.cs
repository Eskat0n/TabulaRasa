namespace TabulaRasa.MetaObjects
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

    internal class Image
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
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent {Cx = 990000L, Cy = 792000L},
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                            {
                                LeftEdge = 0L,
                                TopEdge = 0L,
                                RightEdge = 0L,
                                BottomEdge = 0L
                            },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                            {
                                Id = (UInt32Value) 1U,
                                Name = "Picture"
                            },
                        new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                            new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks {NoChangeAspect = true}),
                        new DocumentFormat.OpenXml.Drawing.Graphic(
                            new DocumentFormat.OpenXml.Drawing.GraphicData(
                                new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                            {
                                                Id = (UInt32Value) 0U,
                                                Name = PictureName
                                            },
                                        new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                        new DocumentFormat.OpenXml.Drawing.Blip(
                                            new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                                new DocumentFormat.OpenXml.Drawing.BlipExtension
                                                    {
                                                        Uri = BlipUriGuid
                                                    })
                                            )
                                            {
                                                Embed = imagePartId,
                                                CompressionState =
                                                    DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                            },
                                        new DocumentFormat.OpenXml.Drawing.Stretch(
                                            new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                    new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                        new DocumentFormat.OpenXml.Drawing.Transform2D(
                                            new DocumentFormat.OpenXml.Drawing.Offset {X = 0L, Y = 0L},
                                            new DocumentFormat.OpenXml.Drawing.Extents {Cx = 990000L, Cy = 792000L}),
                                        new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                            new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                            ) {Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle}))
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