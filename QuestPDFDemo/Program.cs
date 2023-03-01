using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace QuestPDFDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<string> Rows = new List<string>()
            {
                "Goals"
            };
            List<ColumnDef> Columns = new List<ColumnDef>()
            {
                 new ColumnDef{ ColumnName = "", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Days Wrkd", IsBorder = true, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Appts Schd", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Appts Kept", IsBorder = true, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Tel Dials", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Tel Rchs", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Tel Apmt", IsBorder = true, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Meal", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Leads (QS)", IsBorder = false, Heading = "Business Activity"},
                 new ColumnDef{ ColumnName = "Refr Atmp", IsBorder = false, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "Refr Obtn", IsBorder = false, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "New Seens", IsBorder = false, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "New Facts", IsBorder = false, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "Case Opnd", IsBorder = false, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "Clos", IsBorder = true, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "Pts", IsBorder = true, Heading = "Business Activity"},
                   new ColumnDef{ ColumnName = "Core Sis", IsBorder = false, Heading = "Submitted"},
                   new ColumnDef{ ColumnName = "Other Sis", IsBorder = false, Heading = "Submitted"},
                   new ColumnDef{ ColumnName = "New Clnts", IsBorder = false, Heading = "Submitted"},
                   new ColumnDef{ ColumnName = "Core FYC", IsBorder = false, Heading = "Submitted"},
                   new ColumnDef{ ColumnName = "Other FYC", IsBorder = true, Heading = "Submitted"},
                   new ColumnDef{ ColumnName = "Core Sales", IsBorder = false, Heading = "Placed"},
                   new ColumnDef{ ColumnName = "Other Sales", IsBorder = false, Heading = "Placed"},
                   new ColumnDef{ ColumnName = "New Clnts", IsBorder = false, Heading = "Placed"},
                   new ColumnDef{ ColumnName = "Core FYC", IsBorder = false, Heading = "Placed"},
                   new ColumnDef{ ColumnName = "Other FYC", IsBorder = true, Heading = "Placed"},
                   new ColumnDef{ ColumnName = "Total FYC", IsBorder = false, Heading = ""},
            };
            List<ActivityValuesDef> objActivityValues = new List<ActivityValuesDef>()
            {
                new ActivityValuesDef{Heading="Days Worked",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Appmts Scheduled",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Appointment Kept",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Dials",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Reached",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Appmts Made by Phone",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Bussiness Meals",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Qualified Suspect",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Referral Attempts",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="New Seens",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="New Facts",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Cases Opened",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Closing Interviews",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
                new ActivityValuesDef{Heading="Efficiency Points",DollarValue ="NaN", Sale="NaN", Day="NaN", NewClient="Nan" },
            };
            List<KeySkillDef> objKeySkillDef = new List<KeySkillDef>()
            {
                new KeySkillDef{Heading="Dials to Reached",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Reaches to appointment made",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Appointment Efficiency",IsHeading=true},
                new KeySkillDef{Heading="Appointment schedule to Kept % to Reached",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Referral Obtained per attempt",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="QS to New facts %",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="New facts to open cases %",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Cases opened to Closing Interviews %",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Selling Efficiency",IsHeading=true},
                new KeySkillDef{Heading="Cases opened to Total Sales Submitted %",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Average Placed FYC per Core Sales",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Average Placed FYC per Other Sales",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Core Sale Underwriting success",IsHeading=false,Actual="NaN",Goals="25%"},
                new KeySkillDef{Heading="Points per Kept appointment",IsHeading=false,Actual="NaN",Goals="25%"},
            };

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(5, Unit.Millimetre);
                    page.PageColor(Colors.White);
                    page.DefaultTextStyle(x => x.FontSize(8));

                    page.Header().AlignCenter()
                        .Text("Report 1")
                        .SemiBold().FontSize(36).FontColor(Colors.Blue.Medium);

                    uint colCount = Convert.ToUInt32(Columns.Count());
                    page.Content()
                        .PaddingVertical(1, Unit.Centimetre)
                        .Column(x =>
                        {
                            x.Spacing(10);
                            x.Item().Row(row =>
                            {
                                var containerInside = x.Item();
                                containerInside.Table(table =>
                                {
                                    //********Define columns for table
                                    table.ColumnsDefinition(columns =>
                                    {
                                        foreach (var item in Columns)
                                        {
                                            columns.RelativeColumn(string.IsNullOrEmpty(item.ColumnName) ? 40 : 20);
                                        }
                                    });
                                    //*********Add top main header category
                                    table.Header(header =>
                                    {
                                        string mainHeading = "";
                                        foreach (var item in Columns)
                                        {
                                            if (mainHeading != item.Heading)
                                            {
                                                mainHeading = item.Heading;
                                                uint count = Convert.ToUInt32(Columns.Where(col => col.Heading == mainHeading).Count());
                                                header.Cell().ColumnSpan(count).Element(TopHeaderBlock).Text(item.Heading).FontSize(15);

                                            }

                                        }
                                    });
                                    //*********Add table header with all columns and column grouping
                                    table.Header(header =>
                                    {
                                        foreach (var item in Columns)
                                        {
                                            if (item.IsBorder)
                                                header.Cell().Element(GroupHeaderBorderBlock).Element(HeaderBlock).Text(item.ColumnName);
                                            else
                                                header.Cell().Element(HeaderBlock).Text(item.ColumnName);
                                        }
                                    });
                                    //***********add rows with dummy data
                                    uint currrentRow = 1; uint currentColumn = 2;
                                    foreach (var customRow in Rows)
                                    {
                                        currentColumn = 2;
                                        foreach (var col in Columns)
                                        {
                                            if (col.ColumnName == "")
                                                table.Cell().Row(currrentRow).Column(1).Background("#c5d4e9").Text(customRow);
                                            else
                                            {
                                                if (col.IsBorder)
                                                    table.Cell().Row(currrentRow).Column(currentColumn).Element(GroupBorderBlock).Element(DataRowBlock).Text("1.20");
                                                else
                                                    table.Cell().Row(currrentRow).Column(currentColumn).Element(DataRowBlock).Text("1.20");
                                                currentColumn++;
                                            }
                                        }
                                        currrentRow++;
                                    }
                                    //Add Heading 2nd Row
                                    table.Cell().Row(currrentRow).Column(1).Background("#c5d4e9").MinHeight(20).Text("");
                                    table.Cell().Row(currrentRow).Column(2).ColumnSpan(colCount - 1).BorderBottom(1).MinHeight(20).Text("Daily Averages");
                                    currrentRow++;
                                    //Add Nan Row
                                    table.Cell().Row(currrentRow).Column(1).Background("#c5d4e9").Text("");
                                    currentColumn = 2;
                                    foreach (var col in Columns)
                                    {
                                        if (col.ColumnName == "")
                                            table.Cell().Row(currrentRow).Column(1).Text("");
                                        else
                                        {
                                            if (col.IsBorder)
                                                table.Cell().Row(currrentRow).Column(currentColumn).Element(GroupBorderBlock).Element(DataRowBlock).Text("NaN");
                                            else
                                                table.Cell().Row(currrentRow).Column(currentColumn).Element(DataRowBlock).Text("NaN");
                                            currentColumn++;
                                        }
                                    }
                                    currrentRow++;
                                    //Add Total for period Row
                                    table.Cell().Row(currrentRow).Column(1).Background("#c5d4e9").MinHeight(20).Text("");
                                    table.Cell().Row(currrentRow).Column(2).ColumnSpan(colCount - 1).BorderBottom(1).MinHeight(20).Text("Totals for period");

                                    //Add last row
                                    currrentRow++;
                                    table.Cell().Row(currrentRow).Column(1).Background("#c5d4e9").Text("");
                                    currentColumn = 2;
                                    foreach (var col in Columns)
                                    {
                                        if (col.ColumnName == "")
                                            table.Cell().Row(currrentRow).Column(1).Text("");
                                        else
                                        {
                                            if (col.IsBorder)
                                                table.Cell().Row(currrentRow).Column(currentColumn).Element(GroupBorderBlock).Element(DataRowBlock).Text(col.ColumnName == "Pts" ? "0" : "");
                                            else
                                                table.Cell().Row(currrentRow).Column(currentColumn).Element(DataRowBlock).Text("");
                                            currentColumn++;
                                        }
                                    }
                                });
                            });

                            x.Item().Row(row =>
                            {
                                row.RelativeItem(200).Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(120);
                                        columns.ConstantColumn(80);
                                        columns.ConstantColumn(40);
                                        columns.ConstantColumn(40);
                                        columns.ConstantColumn(40);
                                    });
                                    //*********Add top main header category
                                    table.Header(header =>
                                    {
                                        header.Cell().ColumnSpan(5).Element(TopHeaderBlock).Text("Activity Values and Averages").FontSize(15);
                                    });
                                    //*********Add table header with all columns and column grouping
                                    table.Header(header =>
                                    {
                                        header.Cell().Element(HeaderBlock).Text("Activity");
                                        header.Cell().Element(HeaderBlock).Text("Dollar value of each in paid for FYC");
                                        header.Cell().Element(HeaderBlock).Text("Sale");
                                        header.Cell().Element(HeaderBlock).Text("Day");
                                        header.Cell().Element(HeaderBlock).Text("New Client");
                                    });
                                    uint currrentRow = 1;
                                    foreach (var item in objActivityValues)
                                    {
                                        table.Cell().Row(currrentRow).Column(1).Element(ActivityDataRowBlock).Text(item.Heading);
                                        table.Cell().Row(currrentRow).Column(2).Element(ActivityDataRowBlock).Text(item.DollarValue);
                                        table.Cell().Row(currrentRow).Column(3).Element(ActivityDataRowBlock).Text(item.Sale);
                                        table.Cell().Row(currrentRow).Column(4).Element(ActivityDataRowBlock).Text(item.Day);
                                        table.Cell().Row(currrentRow).Column(5).Element(ActivityDataRowBlock).Text(item.NewClient);
                                        currrentRow++;
                                    }
                                    table.Cell().Row(currrentRow).Column(1).ColumnSpan(5).Border(1).Background("#ffffc0").MinHeight(20).Text("*New Leads and new facts are the '10' and the '2' of the famous '10-3-1' new Client Accusation ratio");
                                });
                                row.RelativeItem(200).Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(190);
                                        columns.ConstantColumn(60);
                                        columns.ConstantColumn(60);
                                    });
                                    //*********Add top main header category
                                    table.Header(header =>
                                    {
                                        header.Cell().ColumnSpan(3).Element(TopHeaderBlock).Text("Key Skill Ratios").FontSize(15);
                                    });
                                    //*********Add table header with all columns and column grouping
                                    table.Header(header =>
                                    {
                                        header.Cell().Element(HeaderBlock).Text("Phoning Efficiency");
                                        header.Cell().Element(HeaderBlock).Text("Actual");
                                        header.Cell().Element(HeaderBlock).Text("Goals");
                                    });
                                    uint currrentRow = 1;
                                    foreach (var item in objKeySkillDef)
                                    {
                                        table.Cell().Row(currrentRow).Column(1).Background(item.IsHeading ? "#a7b4c5" : "#fff").Element(ActivityDataRowBlock).Text(item.Heading);
                                        table.Cell().Row(currrentRow).Column(2).Background(item.IsHeading ? "#a7b4c5" : "#fff").Element(ActivityDataRowBlock).Text(item.Actual);
                                        table.Cell().Row(currrentRow).Column(3).Background(item.IsHeading ? "#a7b4c5" : "#fff").Element(ActivityDataRowBlock).Text(item.Goals);
                                        currrentRow++;
                                    }
                                });

                                List<KeySkillDef> objInventoryReportDef = new List<KeySkillDef>()
                                    {
                                        new KeySkillDef{Heading="Short Term Inventory",IsHeading=true},
                                        new KeySkillDef{Heading="# of Cases : 5",IsHeading=false},
                                        new KeySkillDef{Heading="Weighted FYC : $8200",IsHeading=false},
                                        new KeySkillDef{Heading="Long Term Inventory",IsHeading=true},
                                        new KeySkillDef{Heading="# of Cases : 2",IsHeading=false},
                                        new KeySkillDef{Heading="Weighted FYC : $9875",IsHeading=false},
                                        new KeySkillDef{Heading="Submitted Inventory",IsHeading=true},
                                        new KeySkillDef{Heading="# of Cases : 3",IsHeading=false},
                                        new KeySkillDef{Heading="FYC : $4400",IsHeading=false},
                                        new KeySkillDef{Heading="Active Clients",IsHeading=true},
                                        new KeySkillDef{Heading="# of Cases : 0",IsHeading=false},
                                        new KeySkillDef{Heading="FYC : $268500",IsHeading=false},
                                    };
                                row.RelativeItem(100).Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(150);
                                    });
                                    //*********Add top main header category
                                    table.Header(header =>
                                    {
                                        header.Cell().Element(TopHeaderBlock).Text("Inventory report").FontSize(15);
                                    });
                                    //*********Add table header with all columns and column grouping
                                    uint currrentRow = 1;
                                    foreach (var item in objInventoryReportDef)
                                    {
                                        table.Cell().Row(currrentRow).Column(1).MinHeight(item.IsHeading ? 10 : 20).Background(item.IsHeading ? "#a7b4c5" : "#fff").Element(ActivityDataRowBlock).Text(item.Heading);
                                        currrentRow++;
                                    }

                                });

                            });


                        });
                });
            })
            .GeneratePdf(Guid.NewGuid().ToString().Substring(0, 10) + ".pdf");

        }

        static IContainer TopHeaderBlock(IContainer container)
        {
            return container
            .AlignCenter()
            .AlignMiddle();
        }
        static IContainer HeaderBlock(IContainer container)
        {
            return container
            .BorderHorizontal(2)
            .Background("#c5d4e9")
            .ShowOnce()
            .MinWidth(20)
            .MinHeight(30)
            .AlignCenter()
            .AlignMiddle();
        }
        static IContainer DataRowBlock(IContainer container)
        {
            return container
            .BorderBottom(1)
            .ShowOnce()
            .MinWidth(20)
            .MinHeight(20)
            .AlignCenter()
            .AlignMiddle();
        }
        static IContainer GroupHeaderBorderBlock(IContainer container)
        {
            return container
                 .BorderRight(2);

        }
        static IContainer GroupBorderBlock(IContainer container)
        {
            return container
                 .BorderRight(1);

        }
        static IContainer ActivityDataRowBlock(IContainer container)
        {
            return container
            .Border(1)
            .PaddingLeft(5)
            .ShowOnce()
            .MinWidth(20)
            .MinHeight(10)
            .AlignMiddle();
        }
    }
    public class ColumnDef
    {
        public string Heading { get; set; }
        public string ColumnName { get; set; }
        public bool IsBorder { get; set; }
    }
    public class ActivityValuesDef
    {
        public string Heading { get; set; }
        public string DollarValue { get; set; }
        public string Sale { get; set; }
        public string Day { get; set; }
        public string NewClient { get; set; }
    }
    public class KeySkillDef
    {
        public string Heading { get; set; }
        public string Actual { get; set; }
        public string Goals { get; set; }
        public bool IsHeading { get; set; }
    }

}
