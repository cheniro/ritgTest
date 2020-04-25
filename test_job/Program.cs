using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_job
{
    
    public class Tree
    {
        string name;
        Tree parent;
        List<Tree> child = new List<Tree>();
        public Tree(string name, Tree parent = null)
        {
            this.name = name;
            if(parent != null)
            {
                this.parent = parent;
                parent.child.Add(this);
            }
        }
        
        public string GetStruct()
        {
            string ans = "Родителя нет!\n";
            if (parent != null)
                ans = string.Format("Родительская категория: {0}\n", parent.name);
            ans += GetChild(1);
            return ans;
        }

        public string GetChild(int n)
        {
            string ans = printSymbols(n) + string.Format("{0}\n", name);
            if (child.Count > 0)
                foreach (Tree ch in child)
                    ans += ch.GetChild(n + 1);
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
                    break;///////
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
                        Tree newItem;
                        if (i == 0)
                            newItem = new Tree(temp[i]);
                        else
                            newItem = new Tree(temp[i], dict[temp[i - 1]]);

                        keys.Add(temp[i]);
                        dict.Add(temp[i], newItem);
                    }
            }
            string get_c = "";
            do
            {
                Console.WriteLine("");
                for (int i = 0; i < keys.Count; i++)
                    Console.WriteLine(string.Format("{0} - {1}", keys[i], i));

                Console.WriteLine("Выход - q\n");
                get_c = Console.ReadLine();

                if (Int32.TryParse(get_c, out int indx) && indx < keys.Count - 1 && indx >= 0)
                    Console.Write(dict[keys[indx]].GetStruct());
                else if (get_c != "q")
                    Console.WriteLine("Неверныее входные данные!");
            } while (get_c != "q");
        }
    }

    public class ArraySort
    {
        public static void sorting(char[,] letters)
        {
            cw(ref letters);
            int indx_min = -1;
            for (int i = 0; i < letters.GetLength(1) - 1; i++)
            {
                indx_min = ind_min(ref letters, i);
                if (letters[0, indx_min] == letters[0, i] && indx_min != i)
                    indx_min = ind_min2(ref letters, i, indx_min);

                if (indx_min != i)
                {
                    swap(ref letters, i, indx_min);
                    cw(ref letters);
                }
            }
            Console.WriteLine("\nРезультат:\n");
            cw(ref letters);

        }

        static void swap(ref char[,] let, int f, int s)
        {
            int diff = 0;
            for (int i = 0; i < let.GetLength(0); i++)
            {
                diff = let[i, f] - let[i, s];
                let[i, f] = (char)(let[i, f] - diff);
                let[i, s] = (char)(let[i, s] + diff);
            }
        }

        static int ind_min(ref char[,] arr, int st_ind)
        {
            int ind = st_ind;
            for (int i = st_ind + 1; i < arr.GetLength(1); i++)
                if (arr[0, ind] > arr[0, i])
                    ind = i;
                else if (arr[0, ind] == arr[0, i])
                {
                    ind = i;
                    break;
                }
            return ind;
        }

        static int ind_min2(ref char[,] arr, int f, int s)
        {
            int ind = f;
            for (int i = 1; i < arr.GetLength(0); i++)
                if (arr[i, f] > arr[i, s])
                {
                    ind = s;
                    break;
                }
            return ind;
        }

        static void cw(ref char[,] arr)
        {
            for (int t = 0; t < arr.GetLength(0); t++)
            {
                for (int i = 0; i < arr.GetLength(1); i++)
                {
                    Console.Write(arr[t, i] + " ");
                }
                Console.WriteLine("");
            }
            Console.WriteLine("");
        }
    }


    class Program
    {
        static char[,] letters = {{'А', 'П', 'К', 'А', 'Н' },
                                  {'Б', 'З', 'А', 'Б', 'В' },
                                  {'Г', 'П', 'Ы', 'А', 'У' },
                                  {'Х', 'Ш', 'И', 'И', 'Н' }};


        static void Main(string[] args)
        {
            string n_test = "";
            do
            {
                Console.WriteLine("Массив - 1\nКласс - 2\nИначе - Выход");
                n_test = Console.ReadLine();
                if (n_test == "1")
                    ArraySort.sorting(letters);
                else if (n_test == "2")
                    CategoryReader.createTree();
                
            } while (n_test == "1" || n_test == "2");
            Console.WriteLine("Нажмите любую кнопку для выхода.");
            Console.ReadKey();
        }

        

        
    }
}
