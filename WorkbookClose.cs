using System;

namespace Roro.Activities.Excel
{
    public class WorkbookClose : ProcessNodeActivity
    {
        public Input<Text> WorkbookName { get; set; }

        public override void Execute(ActivityContext context)
        {
            var wbName = context.Get(this.WorkbookName);

            ExcelBot.Shared.GetWorkbookByName(wbName, true).Close();

            if (ExcelBot.Shared.GetApp().Workbooks.Count == 0)
            {
                ExcelBot.Shared.Dispose();
            }
        }
    }
}
