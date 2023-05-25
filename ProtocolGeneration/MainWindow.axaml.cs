using System.Collections.Generic;
using Avalonia.Controls;
using ProtocolGeneration.Models;

namespace ProtocolGeneration;

public partial class MainWindow : Window
{
    private GenerateProtocol _generateProtocol;
    public MainWindow()
    {
        InitializeComponent();
        List<Student> students = new();
        _generateProtocol = new GenerateProtocol();
        People mainPeople = new();
        People secondPeople = new();
        List<People> peoples = new();
        _generateProtocol.GenerateFirstProtocol(1, "09.02.07 Информационные системы и программирование", mainPeople, secondPeople, peoples, students, 8, 0, 0 );
    }
}