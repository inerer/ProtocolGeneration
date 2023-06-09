using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Avalonia.Controls;
using Cyriller;
using Cyriller.Model;
using ProtocolGeneration.Models;

namespace ProtocolGeneration;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        var generateProtocol = new GenerateProtocol();

        // _generateProtocol.GenerateFirstProtocol(1, "09.02.07 Информационные системы и программирование", mainPeople, secondPeople, peoples, students, 8, 0, 0, people );
        // _generateProtocol.ThirdProtocol(1, "09.02.07 Информационные системы и программировани", student, mainPeople, secondPeople, peoples, 8, 0, 0 , people);
        // _generateProtocol.FourthProtocol(1,"09.02.07 Информационные системы и программировани", mainPeople, secondPeople, peoples, people, 8, 0, 0 );
        string initialData2 = @"C:\Users\arshi\Downloads\Книга1.csv";
        string initialData = @"C:\Users\arshi\Downloads\Книга2.csv";
        string[] initialDataRows2 = File.ReadAllLines(initialData2, Encoding.UTF8);
        string[] initialDataRows = File.ReadAllLines(initialData,Encoding.UTF8);
        string resultDirectoryPath = Path.Combine(Environment.CurrentDirectory, "output");
        
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


        List<People> peoples = new List<People>();
        foreach (var item in initialDataRows)
        {
            People p = new People()
            {
                FirstName = item.Split(';')[0],
                LastName = item.Split(';')[1],
                MiddleName = item.Split(';')[2],
                Role = item.Split(';')[3]
            };
            peoples.Add(p);
        }

        List<Student> students = new List<Student>();
        
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
                Date = Convert.ToDateTime(item.Split(';')[15]),
                DemoDate = Convert.ToDateTime(item.Split(';')[16])
            };
            students.Add(student);

           
            generateProtocol.ThirdProtocol(1, 
                "09.02.07 Информационные системы и программирование",
                student,
                mainPeople,
                deputyPeople,
                peoples,
                secretaryPeople);
            
            generateProtocol.FifthProtocol(1, 
                "09.02.07 Информационные системы и программирование",
                mainPeople,
                deputyPeople,
                peoples,
                student,
                secretaryPeople
                );
        } 
        
        generateProtocol.GenerateSecondProtocol(1
            ,"09.02.07 Информационные системы и программирование"
            ,mainPeople
            ,deputyPeople
            ,peoples
            ,students
            ,8
            ,0
            ,0
            ,secretaryPeople);
        
        generateProtocol.FourthProtocol(1
            ,"09.02.07 Информационные системы и программирование"
            ,mainPeople
            ,deputyPeople
            ,peoples
            ,secretaryPeople
            ,8
            ,0
            ,0);
        
        
            
           
       
    }
}