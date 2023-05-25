using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Avalonia.Media;
using ProtocolGeneration.Models;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Color = System.Drawing.Color;

namespace ProtocolGeneration;

public class GenerateProtocol
{
    public void GenerateFirstProtocol(int counter, string specialty, People mainPeople, People deputyPeople, List<People> peoples, List<Student>students, int voteYes, int voteNo, int voteMaybe)
    {
        string path = @"C:\Users\arshi\OneDrive\Desktop\Generate\FirstProtocol.docx";
        
        DocX document = DocX.Create(path);
        
        Paragraph paragraph = document.InsertParagraph();
        Paragraph paragraph20 = document.InsertParagraph();
        Paragraph paragraph1 = document.InsertParagraph();
        
        paragraph.AppendLine("Государственное бюджетное профессиональное\nобразовательное учреждение Московской области\n«Серпуховский колледж»")
                .UnderlineColor(System.Drawing.Color.Black)
                .Font("Times New Roman")
                .FontSize(12)
                .Alignment = Alignment.center;
        
        paragraph.AppendLine($"ПРОТОКОЛ № {counter}")
                .Bold()
                .FontSize(12)
                .Font("Times New Roman")
                .Alignment = Alignment.center;
        
        paragraph.AppendLine(
                $"Заседания Государственной экзаменационной комиссии по программе подготовки\nспециалистов среднего звена по специальности {specialty}\nпо приему государственного экзамена в форме демонстрационного экзамена")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman")
            .Alignment = Alignment.center;
      

        paragraph20.AppendLine($"от «19» июня 2023 г.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.left;

        paragraph20.AppendLine("Начало работы ГЭК 7 час. 45 мин.")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph20.AppendLine("Окончание работы ГЭК 19 час. 45 мин.")
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph1.AppendLine("Компетенция ")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman")
            .Alignment=Alignment.left;

        paragraph1.Append($"      09 Программные решения для бизнеса")
            .UnderlineColor(System.Drawing.Color.Black)
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqy")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        
        paragraph1.AppendLine($"Комплект оценочной документации ")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph1.Append("1,2")
            .UnderlineColor(System.Drawing.Color.Black)
            .Font("Times New Roman")
            .FontSize(12);
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqyddddddddddddddd")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph1.AppendLine("Центр проведения демонстрационного экзамена, адрес:")
            .Bold()
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph1.AppendLine("Учебная группа:")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph1.Append("1281")
            .Bold()
            .UnderlineColor(System.Drawing.Color.Black);
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqyddddddddddddd")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph1.AppendLine("Присутствовали:")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph1.AppendLine("Председатель ГЭК  ")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph1.Append($"{mainPeople.FullName}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(System.Drawing.Color.Black);
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqy")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        Paragraph para = document.InsertParagraph();
        
        para.AppendLine("(Ф.И.О., должность)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        Paragraph paragraph3 = document.InsertParagraph();
        
        paragraph3.AppendLine("Зам. председателя: ")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.left;
        
        paragraph3.Append($@"{deputyPeople.FullName}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
        Paragraph paragraph4 = document.InsertParagraph();
        
        paragraph4.AppendLine("(Ф.И.О., должность)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment=Alignment.center;
        
        Paragraph paragraph5 = document.InsertParagraph();
        
        paragraph5.AppendLine("Члены ГЭК: ")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.left;
        
        Paragraph paragraph6 = document.InsertParagraph();
        
        int count = 1;
        foreach (var people in peoples)
        {
            paragraph5.AppendLine($"{count}. {people.FullName}")
                .FontSize(12)
                .Font("Times New Roman");
            
            paragraph6.AppendLine("(Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                .Alignment=Alignment.center;
            
            count++;
        }

        Paragraph paragraph7 = document.InsertParagraph();
        
        paragraph7.AppendLine(
                "Экзаменационная комиссия решила признать, что студент сдал государственный экзамен с оценкой")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman");
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqy")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph10 = document.InsertParagraph();
        paragraph10.AppendLine("Перевод полученного колечиства баллов в оценку")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.right;
        
        Table table = document.AddTable(count, 6);
        table.Rows[0]
            .Cells[0]
            .Paragraphs
            .First()
            .Append("№ п/п")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();

        table.Rows[0]
            .Cells[1]
            .Paragraphs
            .First()
            .Append("Фамилия")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();
        
        table.Rows[0]
            .Cells[2]
            .Paragraphs
            .First()
            .Append("Имя")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();

        table.Rows[0]
            .Cells[3]
            .Paragraphs
            .First()
            .Append("Отчество")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();

        table.Rows[0]
            .Cells[4]
            .Paragraphs
            .First()
            .Append("Итоговые баллы")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();

        table.Rows[0]
            .Cells[5]
            .Paragraphs
            .First()
            .Append("Оценка")
            .Font("Times New Roman")
            .FontSize(12)
            .Bold();

        count = 1;
       
        foreach (var student in students)
        {
            table.Rows[count]
                .Cells[0]
                .Paragraphs
                .First()
                .Append(count.ToString() + ".")
                .Font("Times New Roman")
                .FontSize(12);

            table.Rows[count]
                .Cells[1]
                .Paragraphs
                .First()
                .Append(student.LastName)
                .FontSize(12)
                .Font("Times New Roman");
            
            table.Rows[count]
                .Cells[2]
                .Paragraphs
                .First()
                .Append(student.FirstName)
                .FontSize(12)
                .Font("Times New Roman");
            
            table.Rows[count]
                .Cells[3]
                .Paragraphs
                .First()
                .Append(student.MiddleName)
                .FontSize(12)
                .Font("Times New Roman");
            
            table.Rows[count]
                .Cells[4]
                .Paragraphs
                .First()
                .Append(student.Ball.ToString())
                .FontSize(12)
                .Font("Times New Roman");
            
            table.Rows[count]
                .Cells[5]
                .Paragraphs
                .First()
                .Append(student.Grade.ToString())
                .FontSize(12)
                .Font("Times New Roman");

            count++;
        }
        
        table.Alignment = Alignment.center;
        document.InsertTable(table);

        Paragraph paragraph8 = document.InsertParagraph();
        paragraph8.AppendLine("Итоги голосования председателя и членов аттестационной комиссии:")
            .FontSize(12)
            .Font("Times New Roman")
            .Alignment=Alignment.left;
        
        paragraph8.AppendLine("«За» - ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph8.Append($"{voteYes}")
            .UnderlineColor(Color.Black)
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.Append(" голосов;")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.AppendLine("«Против» - ")
            .Font("Times New Roman")
            .FontSize(12);
       
        paragraph8.Append($"{voteNo}")
            .UnderlineColor(Color.Black)
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.Append(" голосов;")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.AppendLine("«Воздержался» - ")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.Append($"{voteMaybe}")
            .UnderlineColor(Color.Black)
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.Append(" голосов;")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.AppendLine("Решение принято ")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph8.Append(" о")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        
        paragraph8.Append("лвфалjashjwqeihashuwhgsadgwqy")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph9 = document.InsertParagraph();
        
        paragraph9.AppendLine("(Ф.И.О., должность)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment=Alignment.center;
        
        Paragraph paragraph11 = document.InsertParagraph();
        
        paragraph11.AppendLine("Особые мнения членов комиссии")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        
        paragraph11.Append("лвфалjashjwqeihashuwhgsadgwqydjhvjhdjkvhkerjhvkvehurgheirhviuerhiuheiruugviefuefueeueeeeee")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph12 = document.InsertParagraph();

        paragraph12.AppendLine("Председатель ГЭК                ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph12.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph12.Append("         ")
            .Color(Color.White);
        
        paragraph12.Append("SDDDDDDDDDDDD")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("         (подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("                      ")
            .Color(Color.White);

        paragraph13.Append("                ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph14 = document.InsertParagraph();
        
        paragraph14.AppendLine("Главный эксперт ДЭ        ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph14.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph14.Append("      ")
            .Color(Color.White);
        
        paragraph14.Append("SDDDDDDDDDDDD")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph16 = document.InsertParagraph();

        paragraph16.AppendLine("(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph16.Append("                      ")
            .Color(Color.White);

        paragraph16.Append("                ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph15 = document.InsertParagraph();

        paragraph15.AppendLine("Секретарь ГЭК        ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph15.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph15.Append("      ")
            .Color(Color.White);
        
        paragraph15.Append("SDDDDDDDDDDDD")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        Paragraph paragraph17 = document.InsertParagraph();

        paragraph17.AppendLine("(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph17.Append("                      ")
            .Color(Color.White);

        paragraph17.Append("                ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        Process.Start(new ProcessStartInfo
        {
            FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\FirstProtocol.docx",
            UseShellExecute = true
        });

    }
}