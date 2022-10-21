using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace AddAttributesAddin
{
    public class Commands
    {
        [CommandMethod("AddAttribute")]
        public static void AddAttributeReference()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // Attributes definition 
            string Value1 ="AttributeValue1";
            string TAG1 = "TAG1";
            string Value2 = "AttributeValue2";
            string TAG2 = "TAG2";

            // prompts the user to select block references
            var filter = new SelectionFilter(new[] { new TypedValue(0, "INSERT") });
            var psr = ed.GetSelection(filter);
            if (psr.Status != PromptStatus.OK) return;
            dynamic acadApp = Application.AcadApplication;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in psr.Value)
                {
                    // prompts the user to select a block
                    var br = (BlockReference)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                    br.Highlight();
                    // prompts the user to specify the attribute insertion point
                    var ppr = ed.GetPoint("\nSpecify the attribute insertion point: ");
                    br.Unhighlight();
                    // If still no insertion point, continue the request
                    if (ppr.Status != PromptStatus.OK) continue;
                    // Insert attribute value in the insertion point
                    var insPoint = ppr.Value.TransformBy(ed.CurrentUserCoordinateSystem);
                    var attRef1 = new AttributeReference(insPoint, Value1, TAG1, db.Textstyle) { LockPositionInBlock = false };
                    br.AttributeCollection.AppendAttribute(attRef1);
                    tr.AddNewlyCreatedDBObject(attRef1, true);
                    db.TransactionManager.QueueForGraphicsFlush();

                    // prompts the user to specify the second attribute insertion point
                    var ppr2 = ed.GetPoint("\nSpecify the second attribute insertion point: ");
                    br.Unhighlight();
                    // If still no insertion point, continue the request
                    if (ppr2.Status != PromptStatus.OK) continue;
                    var insPoint2 = ppr2.Value.TransformBy(ed.CurrentUserCoordinateSystem);
                    var attRef2 = new AttributeReference(insPoint2, Value2, TAG2, db.Textstyle) { LockPositionInBlock = false };
                    br.AttributeCollection.AppendAttribute(attRef2);
                    tr.AddNewlyCreatedDBObject(attRef2, true);
                    db.TransactionManager.QueueForGraphicsFlush();
                }
                tr.Commit();
            }
        }
    }
}
