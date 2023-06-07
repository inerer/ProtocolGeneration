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
    private readonly GenerateProtocol _generateProtocol;
    public MainWindow()
    {
        InitializeComponent();
        _generateProtocol = new GenerateProtocol();
        
        // _generateProtocol.GenerateFirstProtocol(1, "09.02.07 Информационные системы и программирование", mainPeople, secondPeople, peoples, students, 8, 0, 0, people );
        // _generateProtocol.ThirdProtocol(1, "09.02.07 Информационные системы и программировани", student, mainPeople, secondPeople, peoples, 8, 0, 0 , people);
        // _generateProtocol.FourthProtocol(1,"09.02.07 Информационные системы и программировани", mainPeople, secondPeople, peoples, people, 8, 0, 0 );
        

        string initialData = @"C:\Users\arshi\Downloads\ГЭК.txt";
        string[] initialDataRows = File.ReadAllLines(initialData,Encoding.UTF8);
        string resultDirectoryPath = Path.Combine(Environment.CurrentDirectory, "output");

        List<People> peoples = new List<People>();
        foreach (var item in initialDataRows[2].ToString().Split(';'))
        {
            People p = new People()
            {
                FirstName = item.Split(' ')[0],
                LastName = item.Split(' ')[1],
                MiddleName = item.Split(' ')[2],
                Role = item.Split(',')[1]
            };
            peoples.Add(p);
        }

        List<Student> students = new List<Student>();
        
        foreach (var item in initialDataRows[3].ToString().Split(';'))
        {
            Student p = new Student()
            {
                FirstName = item.Split(',')[0],
                LastName = item.Split(',')[1],
                MiddleName = item.Split(',')[2],
                Ball = decimal.Parse(item.Split(',')[3]),
                CountList = int.Parse(item.Split(',')[4]),
                CountGrap = int.Parse(item.Split(',')[5]),
                Theme = item.Split(',')[6],
                MainTeacher = item.Split(',')[7],
                Opinion = item.Split(',')[8],
                Review = item.Split(',')[9],
                Group = int.Parse(item.Split(',')[10]),
                FirstQuestion = item.Split(',')[11],
                SecondQuestion = item.Split(',')[12],
                ThirdQuestion = item.Split(',')[13],
                SpecialOpinion = item.Split(',')[14],
                Date = Convert.ToDateTime(item.Split(',')[15]),
                DemoDate = Convert.ToDateTime(item.Split(',')[16])
            };
            students.Add(p);
            
            _generateProtocol.ThirdProtocol(1, "09.02.07 Информационные системы и программирование", p, new People(){FirstName = initialDataRows[0].ToString().Split(' ')[0],
                LastName = initialDataRows[0].ToString().Split(' ')[1],MiddleName =  initialDataRows[0].ToString().Split(' ')[2], Role = initialDataRows[0].ToString().Split(';')[1]},
                new People(){FirstName =initialDataRows[1].ToString(),LastName = initialDataRows[1].ToString(),MiddleName = initialDataRows[1].ToString()},
                peoples, int.Parse(initialDataRows[4].Split(';')[0])
                ,int.Parse(initialDataRows[4].Split(';')[1]),int.Parse(initialDataRows[4].Split(';')[2]), new People(){FirstName = "Денис", MiddleName = "Викторович", LastName = "Головин"} );
            
            _generateProtocol.FifthProtocol(1, "09.02.07 Информационные системы и программирование",new People(){FirstName = initialDataRows[0].ToString().Split(' ')[0],
                LastName = initialDataRows[0].ToString().Split(' ')[1],MiddleName =  initialDataRows[0].ToString().Split(' ')[2], Role = initialDataRows[0].ToString().Split(';')[1]}, new People(){FirstName =initialDataRows[1].ToString(),LastName = initialDataRows[1].ToString(),MiddleName = initialDataRows[1].ToString()}, peoples, p,  new People(){FirstName = "Денис", MiddleName = "Викторович", LastName = "Головин"},int.Parse(initialDataRows[4].Split(';')[0])
                ,int.Parse(initialDataRows[4].Split(';')[1]),int.Parse(initialDataRows[4].Split(';')[2]));
        } 
        
        _generateProtocol.GenerateSecondProtocol(1,"09.02.07 Информационные системы и программирование",new People(){FirstName = initialDataRows[0].ToString().Split(' ')[0],
                 LastName = initialDataRows[0].ToString().Split(' ')[1],MiddleName =  initialDataRows[0].ToString().Split(' ')[2], Role = initialDataRows[0].ToString().Split(';')[1]},
                 new People(){FirstName =initialDataRows[1].ToString(),LastName = initialDataRows[1].ToString(),MiddleName = initialDataRows[1].ToString()},
                 peoples,students
                 ,int.Parse(initialDataRows[4].Split(';')[0])
                 ,int.Parse(initialDataRows[4].Split(';')[1]),int.Parse(initialDataRows[4].Split(';')[2]),
                 new People(){FirstName = "Денис", MiddleName = "Викторович", LastName = "Головин"});
        
        _generateProtocol.FourthProtocol(1, "09.02.07 Информационные системы и программирование", new People(){FirstName = initialDataRows[0].ToString().Split(' ')[0],
                   LastName = initialDataRows[0].ToString().Split(' ')[1],MiddleName =  initialDataRows[0].ToString().Split(' ')[2], Role = initialDataRows[0].ToString().Split(';')[1]}, new People(){FirstName = initialDataRows[0].ToString().Split(' ')[0],
            LastName = initialDataRows[0].ToString().Split(' ')[1],MiddleName =  initialDataRows[0].ToString().Split(' ')[2]}, peoples, new People(){FirstName = "Денис", MiddleName = "Викторович", LastName = "Головин"}, int.Parse(initialDataRows[4].Split(';')[0])
            ,int.Parse(initialDataRows[4].Split(';')[1]),int.Parse(initialDataRows[4].Split(';')[2]));
        
        
            
           
       
    }
}