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
                "Daily Plan","Weekly Plan","Monthly Plan","Yearly Plan"
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
                   new ColumnDef{ ColumnName = "Clos", IsBorder = false, Heading = "Business Activity"},
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


                    page.Content()
                        .PaddingVertical(1, Unit.Centimetre)
                        .Column(x =>
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
            .MinHeight(30)
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
    }
    public class ColumnDef
    {
        public string Heading { get; set; }
        public string ColumnName { get; set; }
        public bool IsBorder { get; set; }
    }
}
