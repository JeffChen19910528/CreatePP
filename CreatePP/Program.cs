using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace CreatePP
{
    class Program
    {
        static void Main(string[] args)
        {
            //   animates   the   shapes,   and   runs   the   slide   show.

            Application ppApp = new Application();

            //   Create   a   new   PowerPoint   presentation.
            Presentation objPres = ppApp.Presentations.Add(MsoTriState.msoTrue);

            //   Add   a   slide   to   the   presentation.
            _Slide objSlide = objPres.Slides.Add
            (1, PpSlideLayout.ppLayoutBlank);

            //   Place   two   shapes   on   the   slide.  
            Microsoft.Office.Interop.PowerPoint.Shape objSquareShape = objSlide.Shapes.AddShape
            (MsoAutoShapeType.msoShapeRectangle,
            0, 0, 100, 100);
            Microsoft.Office.Interop.PowerPoint.Shape objTriangleShape = objSlide.Shapes.AddShape
            (MsoAutoShapeType.msoShapeRightTriangle,
            0, 150, 100, 100);

            //   Add   an   animation   sequence.  
            Sequence objSequence =
            objSlide.TimeLine.InteractiveSequences.Add(1);

            //   Add   text   to   the   shapes.  
            objSquareShape.TextFrame.TextRange.Text = "Click   Me! ";
            objTriangleShape.TextFrame.TextRange.Text = "Me   Too! ";



            //   Animate   the   shapes.  
            objSequence.AddEffect(objSquareShape,
            MsoAnimEffect.msoAnimEffectPathStairsDown,
            MsoAnimateByLevel.msoAnimateLevelNone,
            MsoAnimTriggerType.msoAnimTriggerOnShapeClick,
            1);
            objSequence.AddEffect(objTriangleShape,
            MsoAnimEffect.msoAnimEffectPathHorizontalFigure8,
            MsoAnimateByLevel.msoAnimateLevelNone,
            MsoAnimTriggerType.msoAnimTriggerOnShapeClick,
            1);
        }
    }
}
