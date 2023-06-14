using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Cyriller;
using Cyriller.Model;
using ProtocolGeneration.Models;

namespace ProtocolGeneration;

public partial class MainWindow : Window
{
    private string initialData2;
    public MainWindow()
    {
        InitializeComponent();
    }

    private async Task Save()
    {
        var dialog = new OpenFileDialog();
        dialog.Filters?.Add(new FileDialogFilter() 
         
            {Name = "Text", Extensions = {"csv"}});
        var result = await dialog.ShowAsync(this);
        if (result != null)
        {
            Generate(result[0]);
        }
    }

    private async void AddFirstPathButton_OnClick(object? sender, RoutedEventArgs e)
    { 
        try
        {
            await Save();
        }
        catch(Exception c)
        {
         var messageBoxStandardWindow = MessageBox.Avalonia.MessageBoxManager
                .GetMessageBoxStandardWindow("Ошибка", "При генерации произошла ошибка," +
                                                       " скорее всего вы выбрали не подходящий файл");
         await messageBoxStandardWindow.Show();
        }
       
    }

    private void Generate(string data2)
    {
        var generateProtocol = new GenerateProtocol();
        List<Student> students = new List<Student>();
       

        // _generateProtocol.GenerateFirstProtocol(1, "09.02.07 Информационные системы и программирование", mainPeople, secondPeople, peoples, students, 8, 0, 0, people );
        // _generateProtocol.ThirdProtocol(1, "09.02.07 Информационные системы и программировани", student, mainPeople, secondPeople, peoples, 8, 0, 0 , people);
        // _generateProtocol.FourthProtocol(1,"09.02.07 Информационные системы и программировани", mainPeople, secondPeople, peoples, people, 8, 0, 0 );
        initialData2 = data2;
        string[] initialDataRows2 = File.ReadAllLines(initialData2, Encoding.UTF8);
        int count = Convert.ToInt32(initialDataRows2[0].Split(';')[22]);
        int dateStart = Convert.ToInt32(initialDataRows2[0].Split(';')[23]);
       
        People secretaryPeople = new People()
        {
            FirstName = "Денис",
            MiddleName = "Викторович",
            LastName = "Головин"
        };

        People mainPeople = new People()
        {
            FirstName = "Алексей",
            LastName = "Солдатов",
            MiddleName = "Сергеевич",
            Role = "начальник отдела Центра информационных технологий АО «Серпуховский завод «Металлист» "
        };

        People deputyPeople = new People()
        {
            FirstName = "Леонид",
            LastName = "Быковский",
            MiddleName = "Николаевич",
            Role = "заместитель директора по  учебно-производственной работе ГБПОУ МО «Серпуховский колледж»"
        };

        foreach (var item in initialDataRows2)
        {
            Student student = new Student()
            {
                FirstName = item.Split(';')[0],
                LastName = item.Split(';')[1],
                MiddleName = item.Split(';')[2],
                Ball = decimal.Parse(item.Split(';')[3]),
                CountList = int.Parse(item.Split(';')[4]),
                CountGrap = int.Parse(item.Split(';')[5]),
                Theme = item.Split(';')[6],
                MainTeacher = item.Split(';')[7],
                Opinion = item.Split(';')[8],
                Review = item.Split(';')[9],
                Group = int.Parse(item.Split(';')[10]),
                FirstQuestion = item.Split(';')[11],
                SecondQuestion = item.Split(';')[12],
                ThirdQuestion = item.Split(';')[13],
                SpecialOpinion = item.Split(';')[14],
                Date = DateOnly.Parse(item.Split(';')[15]),
                DemoDate = DateOnly.Parse(item.Split(';')[16]),
                VKRGrade = int.Parse(item.Split(';')[17]),
                voteYes = int.Parse(item.Split(';')[18]),
                voteMaybe = int.Parse(item.Split(';')[19]),
                voteNo = int.Parse(item.Split(';')[20]),
                DiplomCathegory = item.Split(';')[21]
            };
            students.Add(student);

            count++;
            generateProtocol.ThirdProtocol(count, 
                "09.02.07 Информационные системы и программирование",
                student,
                mainPeople,
                deputyPeople,
                Peoples(),
                secretaryPeople,
                dateStart);
            count++;
            
                // generateProtocol.FifthProtocol(count, 
                // "09.02.07 Информационные системы и программирование",
                // mainPeople,
                // deputyPeople,
                // Peoples(),
                // student,
                // secretaryPeople,
                // dateStart);
        } 
        count = Convert.ToInt32(initialDataRows2[0].Split(';')[22]);
        // generateProtocol.GenerateSecondProtocol(count
        //     ,"09.02.07 Информационные системы и программирование"
        //     ,mainPeople
        //     ,deputyPeople
        //     ,Peoples()
        //     ,students
        //     ,8
        //     ,0
        //     ,0
        //     ,secretaryPeople,
        //     dateStart);
        
        // generateProtocol.FourthProtocol(1
        //     ,"09.02.07 Информационные системы и программирование"
        //     ,mainPeople
        //     ,deputyPeople
        //     ,Peoples()
        //     ,secretaryPeople
        //     ,8
        //     ,0
        //     ,0,
        //     dateStart);
        var messageBoxStandardWindow = MessageBox.Avalonia.MessageBoxManager
            .GetMessageBoxStandardWindow("Поздравляю!", "Протоколы сгенерированы!");
        messageBoxStandardWindow.Show();
    }

    private List<People> Peoples()
    {
        List<People> peoples = new List<People>()
        {
            new People()
            {
                LastName = "Черникова",
                FirstName = "Лилия",
                MiddleName = "Валентиновка",
                Role = "Преподаватель ГБПОУ МО «Серпуховский колледж»"
            },
            new People()
            {
                LastName = "Головин",
                FirstName = "Денис",
                MiddleName = "Викторович",
                Role = "Преподаватель ГБПОУ МО «Серпуховский колледж»"
            },
            new People()
            {
                LastName = "Бурцев",
                FirstName = "Павел",
                MiddleName = "Константинович",
                Role = "Преподаватель ГБПОУ МО «Серпуховский колледж»"
            },
            new People()
            {
                LastName = "Золотухина",
                FirstName = "Ирина",
                MiddleName = "Игоревна",
                Role = "Специалист поддержки, ООО«Авито Тех»"
            },
            new People()
            {
                LastName = "Кривцов",
                FirstName = "Павел",
                MiddleName = "Николаевич",
                Role = "старший, преподаватель кафедры информационных технологий Филиал «Протвино» государственного" +
                       " университета «Дубна»"
            },
            new People()
            {
                LastName = "Архипова",
                FirstName = "Светлана",
                MiddleName = "Станиславовна",
                Role = "старший специалист АНО «Институт инженерной физики»"
            },
            new People()
            {
                LastName = "Фаст",
                FirstName = "Владимир",
                MiddleName = "Владимирович",
                Role = "инженер-программист, ИП Юсупов Б.И."
            }
        };
        return peoples;
    }
}