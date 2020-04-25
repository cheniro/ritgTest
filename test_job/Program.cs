using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_job
{
    class Program
    {
        static char[,] letters = {{'А', 'П', 'К', 'А', 'Н' },
                                  {'Б', 'З', 'А', 'Б', 'В' },
                                  {'Г', 'П', 'Ы', 'А', 'У' },
                                  {'Х', 'Ш', 'И', 'И', 'Н' }};

        static void Main(string[] args)
        {
            string option = "";
            bool exit = false;
            while (!exit)
            {
                Console.WriteLine("Выберите опцию:");
                Console.WriteLine("Задание с массивом - 1\n" +
                                  "Задание с классом - 2\n" +
                                  "Выход - Любой другой символ");
                option = Console.ReadLine();

                switch (option)
                {
                    case "1": ArraySort.sorting(letters); break;
                    case "2": CategoryReader.createTree(); break;
                    default: exit = true; break;
                }

            }
        }
    }

    public class Tree
    {
        string name;
        Tree parent;
        List<Tree> child = new List<Tree>();
        public Tree(string name)
        {
            this.name = name;
        }
        
        public void addParent(Tree parent)
        {
            this.parent = parent;
            parent.child.Add(this);
        }

        public string getStruct()
        {
            string ans = "Родителя нет!\n";
            if (parent != null)
                ans = string.Format("Родительская категория: {0}\n", parent.name);
            ans += getChild(1);
            return ans;
        }

        public string getChild(int n)
        {
            string ans = printSymbols(n) + string.Format("{0}\n", name);
            if (child.Count > 0)
                foreach (Tree ch in child)
                    ans += ch.getChild(n + 1);
            return ans;
        }

        string printSymbols(int n)
        {
            string ans = "-";
            for (int i = 1; i < n; i++)
                ans += "--";
            return ans;
        }
    }
    public class CategoryReader
    {
        public static void createTree()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Excel.Workbook workbook = excel.Workbooks.Open(path + @"/categories.csv");
            Excel.Worksheet sheet = workbook.Worksheets[1];

            List<string> data = new List<string>();
            string rowValue = "";
            for (int i = 2; i <= sheet.Rows.Count; i++)
            {
                rowValue = sheet.Cells[i, 1].Value;
                if (rowValue != null)
                    data.Add(rowValue);
                else
                    break;
            }

            workbook.Close(false);
            excel.Application.Quit();

            Dictionary<string, Tree> dict = new Dictionary<string, Tree>();
            List<string> keys = new List<string>();
            string[] temp = null;
            foreach (string row in data)
            {
                temp = row.Split(';');
                for (int i = 0; i < temp.Length; i++)
                    if (temp[i] != "" && !dict.ContainsKey(temp[i]))
                    {
                        Tree newItem = new Tree(temp[i]);

                        if (i > 0)
                            newItem.addParent(dict[temp[i - 1]]);

                        keys.Add(temp[i]);
                        dict.Add(temp[i], newItem);
                    }
            }
            
            string selectVal = "";
            do
            {
                Console.WriteLine("\nВыберите категорию:");
                for (int i = 0; i < keys.Count; i++)
                    Console.WriteLine(string.Format("{0} - {1}", keys[i], i));

                Console.WriteLine("Выход - q\n");
                selectVal = Console.ReadLine();

                if (Int32.TryParse(selectVal, out int index) && index < keys.Count && index >= 0)
                    Console.Write(dict[keys[index]].getStruct());
                else if (selectVal != "q")
                    Console.WriteLine("Неверныее входные данные!");
            } while (selectVal != "q");
        }
    }

    public class ArraySort
    {
        public static void sorting(char[,] letters)
        {
            printArray(ref letters);
            int indexMin = -1;
            for (int i = 0; i < letters.GetLength(1) - 1; i++)
            {
                indexMin = findIndexMinColumn(ref letters, i);
                if (letters[0, indexMin] == letters[0, i] && indexMin != i)
                    indexMin = findIndexMinValue(ref letters, i, indexMin);

                if (indexMin != i)
                    swap(ref letters, i, indexMin);
            }
            Console.WriteLine("\nРезультат:\n");
            printArray(ref letters);

        }

        static void swap(ref char[,] letters, int first, int second)
        {
            int diff = 0;
            for (int i = 0; i < letters.GetLength(0); i++)
            {
                diff = letters[i, first] - letters[i, second];
                letters[i, first] = (char)(letters[i, first] - diff);
                letters[i, second] = (char)(letters[i, second] + diff);
            }
        }

        static int findIndexMinColumn(ref char[,] letters, int startIndex)
        {
            int index = startIndex;
            for (int i = startIndex + 1; i < letters.GetLength(1); i++)
                if (letters[0, index] > letters[0, i])
                    index = i;
                else if (letters[0, index] == letters[0, i])
                {
                    index = i;
                    break;
                }
            return index;
        }

        static int findIndexMinValue(ref char[,] letters, int first, int second)
        {
            int index = first;
            for (int i = 1; i < letters.GetLength(0); i++)
                if (letters[i, first] > letters[i, second])
                {
                    index = second;
                    break;
                }
            return index;
        }

        static void printArray(ref char[,] letters)
        {
            for (int t = 0; t < letters.GetLength(0); t++)
            {
                for (int i = 0; i < letters.GetLength(1); i++)
                    Console.Write(letters[t, i] + " ");
                Console.WriteLine("");
            }
            Console.WriteLine("");
        }
    }


    
}
