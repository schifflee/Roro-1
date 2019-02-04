using System;

namespace Roro.Activities.Excel
{
    public class WorksheetCreate : ProcessNodeActivity
    {
        public Input<Text> WorkbookName { get; set; }

        public Input<Text> WorksheetName { get; set; }

        public override void Execute(ActivityContext context)
        {
            var wbName = context.Get(this.WorkbookName);
            var wsName = context.Get(this.WorksheetName);

            ExcelBot.Shared.GetWorkbookByName(wbName, true).Worksheets.Add(wsName);
        }
    }
}
