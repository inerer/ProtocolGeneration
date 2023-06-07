using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Avalonia.Media;
using Cyriller;
using Cyriller.Model;
using ProtocolGeneration.Models;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Color = System.Drawing.Color;

namespace ProtocolGeneration;

public class GenerateProtocol
{
    public void GenerateFirstProtocol(int counter, string specialty, People mainPeople, People deputyPeople, List<People> peoples, List<Student>students, int voteYes, int voteNo, int voteMaybe, People secretaryPeople)
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
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqy")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph1.AppendLine("\t\t\t\t\t\t Государственное бюджетное\n\t\t\t\t\t\tпрофессиональное образовательное \n\t\t\t\t\t\tучреждение Московской области\n\t\t\t\t\t\t«Серпуховский колледж», Московская\n\t\t\t\t\t\t область, г. Серпухов, пос. Большевик,\n \t\t\t\t\t\t ")
            .Bold()
            .Font("Times New Roman")
            .FontSize(12);
        paragraph1.Append("ул.Ленина, д.52")
            .Bold()
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        paragraph1.Append("\nЦентр проведения демонстрационного\nэкзамена, адрес:")
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
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .Alignment = Alignment.left;
            paragraph5.Append("psdaodsadasjdkas")
                .Color(Color.White)
                .UnderlineColor(Color.Black);
            
            paragraph5.Append("\n\t\t\t\t\t             (Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                ;
            
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
        
        Table table = document.AddTable(students.Count + 1, 6);
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
        
        paragraph11.Append("лвфалjashjdfdfrhrrigigijtihjtijtj")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph11.AppendLine("")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        paragraph11.InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 1, Color.Black);
        
        Paragraph paragraph12 = document.InsertParagraph();

        paragraph12.AppendLine("Председатель ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph12.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph12.Append("\t")
            .Color(Color.White);

        FIOGenerate(paragraph12, mainPeople);
        
        // paragraph12.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("\t")
            .Color(Color.White);

        paragraph13.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph14 = document.InsertParagraph();
        
        paragraph14.AppendLine("Главный эксперт ДЭ \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph14.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph14.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph14,deputyPeople);
        // paragraph14.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph16 = document.InsertParagraph();

        paragraph16.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph16.Append("\t")
            .Color(Color.White);

        paragraph16.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph15 = document.InsertParagraph();

        paragraph15.AppendLine("Секретарь ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph15.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph15.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph15, secretaryPeople);
        // paragraph15.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph17 = document.InsertParagraph();

        paragraph17.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph17.Append("                      ")
            .Color(Color.White);

        paragraph17.Append("\t ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        // Process.Start(new ProcessStartInfo
        // {
        //     FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\FirstProtocol.docx",
        //     UseShellExecute = true
        // });

    }

    public void FIOGenerate(Paragraph paragraph, People people)
    {
        int maxLenght = 22;

        if (people.FIO != null)
        {
            int fioLength = people.FIO.Length;
        
            paragraph.Append($"{people.FIO}")
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .FontSize(12);
        
            maxLenght -= fioLength;
        }

        for (int i = 0; i < maxLenght; i++)
        {
            paragraph.Append("S")
                .UnderlineColor(Color.Black)
                .Color(Color.White);
        }
        
    } 

    public void GenerateSecondProtocol(int counter, string specialty, People mainPeople, People deputyPeople, List<People> peoples, List<Student>students, int voteYes, int voteNo, int voteMaybe, People secretaryPeople)
    {
         string path = @"C:\Users\arshi\OneDrive\Desktop\Generate\По_переводу.docx";
        
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
                $"Заседания Государственной экзаменационной комиссии по программе подготовки\nспециалистов среднего звена по специальности {specialty}\nпо переводу балллов демонстрационного экзамена в оценку")
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
        paragraph1.Append("лвфалjashjwqeihashuwhgsadgwqy")
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
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .Alignment = Alignment.left;
            paragraph5.Append("psdaodsadasjdkas")
                .Color(Color.White)
                .UnderlineColor(Color.Black);
            
            paragraph5.Append("\n\t\t\t\t\t             (Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                ;
            
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
        
        Table table = document.AddTable(students.Count+1, 6);
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
        
        paragraph11.Append("лвфалjashjdfdfrhrrigigijtihjtijtj")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph11.AppendLine("")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        paragraph11.InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 1, Color.Black);
        
        Paragraph paragraph12 = document.InsertParagraph();

        paragraph12.AppendLine("Председатель ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph12.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph12.Append("\t")
            .Color(Color.White);

        FIOGenerate(paragraph12, mainPeople);
        
        // paragraph12.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("\t")
            .Color(Color.White);

        paragraph13.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph14 = document.InsertParagraph();
        
        paragraph14.AppendLine("Главный эксперт ДЭ \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph14.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph14.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph14,deputyPeople);
        // paragraph14.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph16 = document.InsertParagraph();

        paragraph16.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph16.Append("\t")
            .Color(Color.White);

        paragraph16.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        Paragraph paragraph15 = document.InsertParagraph();

        paragraph15.AppendLine("Секретарь ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph15.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph15.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph15, secretaryPeople);
        // paragraph15.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph17 = document.InsertParagraph();

        paragraph17.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph17.Append("                      ")
            .Color(Color.White);

        paragraph17.Append("\t ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        // Process.Start(new ProcessStartInfo
        // {
        //     FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\SecondProtocol.docx",
        //     UseShellExecute = true
        // });

    }

    public void ThirdProtocol(int counter, string specialty, Student student, People mainPeople, People deputyPeople, List<People> peoples, int voteYes, int voteNo, int voteMaybe, People secretaryPeople)
    {
        CyrName cyrName = new Cyriller.CyrName();

        string path = $@"C:\Users\arshi\OneDrive\Desktop\Generate\{student.FullName}_Протокол защиты.docx";
        
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
                $"Заседания Государственной экзаменационной комиссии по программе подготовки\nспециалистов среднего звена по специальности {specialty}\nпо рассмотрению выпускной квалификационной работы")
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
        
        paragraph20.AppendLine($"По рассмотрению выпускной квалификационной работы студен {cyrName.Decline(student.FullName, CasesEnum.Genitive).ToString()}")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph20.AppendLine($"На тему \t{student.Theme}")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph20.AppendLine($"Выпускная квалификационная работы выполнена под руководством\n{student.MainTeacher}");
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
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .Alignment = Alignment.left;
            paragraph5.Append("psdaodsadasjdkas")
                .Color(Color.White)
                .UnderlineColor(Color.Black);
            
            paragraph5.Append("\n\t\t\t\t\t             (Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                ;
            
            count++;
        }

        paragraph5.AppendLine("В ГЭК представлены следующие материалы")
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph5.AppendLine("1. Сводная ведомость о сданных студентом ")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph5.Append($"{cyrName.Decline(student.FullName, CasesEnum.Genitive).ToString()}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
        paragraph5.Append("\n\t\t\t\t\t\t\tФ.И.О")
            .FontSize(6);
        
        paragraph5.AppendLine("экзаменах и зачетах и о выполнении им учебного плана.")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.AppendLine("2. Текст выпускной квалификационной работы на ")
            .Font("Times New Roman")
            .FontSize(12);
        
        paragraph5.Append("выф")
            .Color(Color.White)
            .UnderlineColor(Color.Black);

        paragraph5.Append($"{student.CountList}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
        paragraph5.Append("выф")
            .Color(Color.White)
            .UnderlineColor(Color.Black);

        paragraph5.Append("страницах")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.AppendLine("Чертежи, схемы, графики к ВКР на ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.Append($"{student.CountGrap}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("4. Отзыв руководителя ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.Opinion}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("5. Рецензия")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.Review}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("Слушали защиту выпускной квалификационной работы студентом ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.Append("4")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);

        paragraph5.Append("курса")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.AppendLine($"{student.Group}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("группы ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.FullName} ")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n\t\tФ.И.О.")
            .FontSize(6)
            .Font("Times New Roman");

        paragraph5.AppendLine("После сообщения о выполненной ВКР в течении ")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.Append("5")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append(" минут студенту(ке) были заданы следуюшие вопросы:")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.AppendLine($"1. {student.FirstQuestion}")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine($"2. {student.SecondQuestion}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine($"3. {student.ThirdQuestion}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("Общая характеристика ответов студента(ки) на вопросы:")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"\n {student.SpecialOpinion}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
        paragraph5.AppendLine("Государственная экзамеционная коммисия постановляет:")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"\n{student.FullName}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n\t\t (Ф.И.О студента)")
            .FontSize(6)
            .Font("Times New Roman");

        paragraph5.AppendLine("выполнил и защитил выпускную работу с оценкой ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.Grade}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
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
        
        paragraph11.Append("лвфалjashjdfdfrhrrigigijtihjtijtj")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph11.AppendLine("")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        
        paragraph11.InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 1, Color.Black);
        
        Paragraph paragraph12 = document.InsertParagraph();

        paragraph12.AppendLine("Председатель ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph12.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph12.Append("\t")
            .Color(Color.White);

        FIOGenerate(paragraph12, mainPeople);
        
        // paragraph12.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("\t")
            .Color(Color.White);

        paragraph13.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");

        Paragraph paragraph15 = document.InsertParagraph();

        paragraph15.AppendLine("Секретарь ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph15.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph15.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph15, secretaryPeople);
        // paragraph15.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph17 = document.InsertParagraph();

        paragraph17.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph17.Append("                      ")
            .Color(Color.White);

        paragraph17.Append("\t ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        // Process.Start(new ProcessStartInfo
        // {
        //     FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\ThirdProtocol.docx",
        //     UseShellExecute = true
        // });
    }

    public void FourthProtocol(int counter, string specialty, People mainPeople, People deputyPeople, List<People> peoples, People secrataryPeople, int voteYes, int voteNo, int voteMaybe)
    {
        CyrName cyrName = new Cyriller.CyrName();
        
        string path = @"C:\Users\arshi\OneDrive\Desktop\Generate\Об_избрании_председателя.docx";
        
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
                $"Заседания Государственной экзаменационной комиссии по программе подготовки\nспециалистов среднего звена по специальности {specialty}\nоб избрании секретаря")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman")
            .Alignment = Alignment.center;
        
        paragraph20.AppendLine($"от «19» июня 2023 г.")
            .Font("Times New Roman")
            .FontSize(12)
            .Alignment = Alignment.left;

        paragraph20.AppendLine("Начало работы ГЭК 8 час. 45 мин.")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph20.AppendLine("Окончание работы ГЭК 10 час. 00 мин.")
            .FontSize(12)
            .Font("Times New Roman");
        
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
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .Alignment = Alignment.left;
            paragraph5.Append("psdaodsadasjdkas")
                .Color(Color.White)
                .UnderlineColor(Color.Black);
            
            paragraph5.Append("\n\t\t\t\t\t             (Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                ;
            
            count++;
        }

        paragraph5.AppendLine("Повестка:")
            .Bold()
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.AppendLine(
                $"\t Избрание секретаря ГЭК.\n По первому вопросу слушали председателя ГЭК ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{cyrName.Decline(mainPeople.FIO, CasesEnum.Genitive).ToString()} ")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
        paragraph5.Append("который")
            .FontSize(12)
            .Font("Times New Roman");
        
        paragraph5.Append("\n\t\t\t\t\t\t\t\tФ.И.О")
            .Font("Times New Roman")
            .FontSize(6);
        
        paragraph5.AppendLine("предложил избрать секретарем ГЭК ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"Головина Д.В.")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n\t\t\t\t\tФ.И.О")
            .FontSize(6)
            .Font("Times New Roman");
        
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
        
        paragraph8.Append(" голосов.")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph8.AppendLine("Решили: ")
            .Bold()
            .FontSize(12)
            .Font("Times New Roman");

        paragraph8.Append("избрать секретарем ГЭК ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph8.Append($"Головина Д.В.")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph8.AppendLine("Особые мнения членов комиссии    ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph8.Append("фывыфвцйуйцвыфвыфвцйвцй")
            .FontSize(12)
            .Font("Times New Roman")
            .Color(Color.White)
            .UnderlineColor(Color.Black);

        paragraph8.AppendLine();
        
        paragraph8.AppendLine("Председатель ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph8.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph8.Append("\t")
            .Color(Color.White);

        FIOGenerate(paragraph8, mainPeople);
        
        // paragraph12.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("\t")
            .Color(Color.White);

        paragraph13.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        // Process.Start(new ProcessStartInfo
        // {
        //     FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\FourthProtocol.docx",
        //     UseShellExecute = true
        // });

    }

    public void FifthProtocol(int counter,string specialty, People mainPeople, People deputyPeople, List<People> peoples, Student student, People secretaryPeople, int voteYes, int voteMaybe, int voteNo)
    {
        CyrName cyrName = new Cyriller.CyrName();
        
        string path = $@"C:\Users\arshi\OneDrive\Desktop\Generate\{student.FullName}_о присвоении.docx";
        
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
                $"Заседания Государственной экзаменационной комиссии по программе подготовки\nспециалистов среднего звена по специальности {specialty}\nО результатах сдачи государственной итоговой аттестации, присвоенииквалификации студенту по результатам государственной итоговой аттестации и выдаче диплома о среднем профессиональном образовании")
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
                .UnderlineColor(Color.Black)
                .Font("Times New Roman")
                .Alignment = Alignment.left;
            paragraph5.Append("psdaodsadasjdkas")
                .Color(Color.White)
                .UnderlineColor(Color.Black);
            
            paragraph5.Append("\n\t\t\t\t\t             (Ф.И.О., должность)")
                .Font("Times New Roman")
                .FontSize(6)
                ;
            
            count++;
        }

        paragraph5.AppendLine("Студент(ка)")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.FullName}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n\t\t\tФ.И.О.")
            .FontSize(6)
            .Font("Times New Roman");

        paragraph5.AppendLine("Защитил(а) выпускную квалификационную работу (ВКР) с оценкой ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.VKRGrade}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n\t\t\t\t (цифрой,прописью)")
            .Font("Times New Roman")
            .FontSize(6);

        paragraph5.AppendLine($"{student.Date}")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append("и сдал(а) демонстрационный экзамен с оценкой ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.WordGrade}")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append("\n(дата защиты ВКР) \t\t\t\t\t\t(цифрой,прописью)")
            .FontSize(6)
            .Font("Times New Roman");

        paragraph5.AppendLine($"{student.DemoDate}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.Append("\n(дата сдачи ДЭ)")
            .FontSize(6)
            .Font("Times New Roman");

        paragraph5.AppendLine("Государственная экзаменационная комиссия постановляет: ")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.AppendLine("Присвоить студенту(ке)\t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph5.Append($"{cyrName.Decline(student.FullName, CasesEnum.Dative)}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("квалификацию\t")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append("Программист")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("по специальности\t")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append("09.02.07 Информационные системы и программирование")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);

        paragraph5.AppendLine("Выдать диплом о среднем профессиональном образовании")
            .FontSize(12)
            .Font("Times New Roman");

        paragraph5.Append($"{student.DiplomCathegory}")
            .FontSize(12)
            .Font("Times New Roman")
            .UnderlineColor(Color.Black);
        
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
        
        paragraph11.Append("лвфалjashjdfdfrhrrigigijtihjtijtj")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph11.AppendLine("")
            .Font("Times New Roman")
            .FontSize(12)
            .UnderlineColor(Color.Black);
        
        paragraph11.InsertHorizontalLine(HorizontalBorderPosition.bottom, BorderStyle.Tcbs_single, 6, 1, Color.Black);

        Paragraph paragraph12 = document.InsertParagraph();

        paragraph12.AppendLine("Председатель ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph12.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph12.Append("\t")
            .Color(Color.White);

        FIOGenerate(paragraph12, mainPeople);
        
        // paragraph12.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph13 = document.InsertParagraph();

        paragraph13.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph13.Append("\t")
            .Color(Color.White);

        paragraph13.Append("\t \tФИО")
            .FontSize(6)
            .Font("Times New Roman");

        Paragraph paragraph15 = document.InsertParagraph();

        paragraph15.AppendLine("Секретарь ГЭК\t \t \t \t")
            .Font("Times New Roman")
            .FontSize(12);

        paragraph15.Append("SDDDDDDDDDDDD    ")
            .UnderlineColor(Color.Black)
            .Color(Color.White);
        
        paragraph15.Append("\t")
            .Color(Color.White);
        
        FIOGenerate(paragraph15, secretaryPeople);
        // paragraph15.Append("SDDDDDDDDDDDD")
        //     .UnderlineColor(Color.Black)
        //     .Color(Color.White);
        
        Paragraph paragraph17 = document.InsertParagraph();

        paragraph17.AppendLine("\t \t \t \t \t(подпись)")
            .Font("Times New Roman")
            .FontSize(6)
            .Alignment = Alignment.center;
        
        paragraph17.Append("                      ")
            .Color(Color.White);

        paragraph17.Append("\t ФИО")
            .FontSize(6)
            .Font("Times New Roman");
        
        document.Save();
        // Process.Start(new ProcessStartInfo
        // {
        //     FileName =  @"C:\Users\arshi\OneDrive\Desktop\Generate\FifthProtocol.docx",
        //     UseShellExecute = true
        // });
        
    }
}