
using System;
using System.Drawing;
using System.Linq;
using System.Text;
using Spire.Presentation;
using Spire.Presentation.Drawing;

class Program 
{
    static void Main()
    {
    
        //Load the sample PowerPoint file
        Presentation Currentresentation = new Presentation();
        Currentresentation.LoadFromFile(@"auxi.pptx");


        //Declare variables 
        ISlide slide;
        IAutoShape shape;
        TextParagraph paragraph;
        TextRange textRange;
        int textcounter = 0;



        //Loop through the slides
        for (int i = 0; i < Currentresentation.Slides.Count; i++)
        {

            //Get the specific slide
            slide = Currentresentation.Slides[i];
            //Loop through the shapes in the slide
            for (int j = 0; j < slide.Shapes.Count; j++)
            {
                //Determine if the shape is an IAutoshape object
                if (slide.Shapes[j] is IAutoShape)
                {
                    //Get the specific shape
                    shape = (IAutoShape)slide.Shapes[j];
                    //Loop through the paragraphs in the shape
                    for (int k = 0; k < shape.TextFrame.Paragraphs.Count; k++)
                    {
                        //Get the specific paragraph
                        paragraph = shape.TextFrame.Paragraphs[k];

                        //Loop through the text ranges in the paragraph
                        for (int m = 0; m < paragraph.TextRanges.Count; m++)
                        {

                            //Get the specific text range
                            textRange = paragraph.TextRanges[m];
                            string text = textRange.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                textcounter++;
                                if (textcounter == 1)
                                {
                                    textRange.Format.LatinFont = new TextFont("Calibri (Body)");
                                    //Change text color
                                    textRange.Format.Fill.FillType = FillFormatType.Solid;
                                    textRange.Format.Fill.SolidColor.KnownColor = KnownColors.Black;
                                    //Make text bold and 
                                    shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Center;

                                }
                           

                            }
                        
                        }
                    }

                }

            }

        }


    
        Presentation presentation = new Presentation();
        // Another way :) 
        Presentation ppt = new Presentation();
        ppt.LoadFromFile(@"auxi.pptx");
        //define the source slide and target slide
        ISlide sourceSlide = ppt.Slides[1];
        ISlide targetSlide = presentation.Slides[0];
        //copy the second shape from the source slide to the target slide

        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[0]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[1]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[2]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[3]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[4]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[5]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[6]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[7]);
        targetSlide.Shapes.AddShape((Shape)sourceSlide.Shapes[8]);
      
        //save the document to file
        presentation.SaveToFile(@"auxiOutput.pptx", FileFormat.Pptx2019);
        DeleteSlide(@"auxiOutput.pptx", 0);

        Console.WriteLine("------------------------------------------------------------");
        Console.WriteLine("------------Presentation has been repolished----------------");
        Console.WriteLine("------------------------------------------------------------");
    }

    public static void DeleteSlide(string presentationFile, int slideIndex)
    {
        //Get PPT document
        Presentation presentation = new Presentation();
        presentation.LoadFromFile(presentationFile);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentOutOfRangeException("slideIndex");
        }
        //Remove the slide by slide index
        presentation.Slides.RemoveAt(slideIndex);

        //Save PPT document
        presentation.SaveToFile(@"..\..\Documents\result.pptx", FileFormat.Pptx2010);
    }

}
