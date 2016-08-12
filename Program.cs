using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace GenerateCV
{     
    static class Program
    {       

        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {            
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FeedFields());
        
        }

        #region comment
        /*
        const int h= 12;
          
            impAbs imp = new impAbs();
          // var tsk= Task.Factory.StartNew(() => { DoSomeWork(); DoSomeOtherWork(); });
            Console.WriteLine("after");
            int[] t = { 1,2,3 };
            string[] t2 = { "aaaa","bbbb","cccc"};
            char[] t3 = { 'a','b','c'};
            string t4 = "rien";
            imp.testTab(t);
            imp.testTab(t2);
            imp.testTab(t3);
            imp.testTab(t4,0);
           
          // tsk.Wait();

       //     Parallel.Invoke(() => DoSomeWork(), () => DoSomeOtherWork());

            var rr=t2.Select(x => x.ToLower());

            var mm= rr.ToLookup(x => x, x => x + " !!");

            int aa = 1;
            int bb = 2;
            if ((aa-- == 0) && (--bb == 0))
            { }

                imp.faireklk();

            String s1 = "12";
            String s2 = "12";
            Int32 o = 1;
            Int32 oo = 1;


            var res1 = s1.Equals(s2);
            var res2 = s1 == s2;

            res1 = o.Equals(oo);
            res2 = o == oo;


            var tt = 15;
            Console.WriteLine(checked(2147483647 + tt));
            impAbs gg = new impAbs();
           
            

            var len = 11;
            Fibonacci_Recursive(len);
            Console.WriteLine();
            int a = 0, b = 1, c = 0;
            Console.Write("{0} {1}", a, b);

            for (int i = 2; i < len; i++)
            {
                c = a + b;
                Console.Write(" {0}", c);
                a = b;
                b = c;
            }

           


            DateTime db = new DateTime(1984, 05, 19);
            DateTime now = DateTime.Now;

            var res = (now-db).Days/364;

            //Pyramid(8);
            //double[] ts = {-1,-1.5,0.23,-0.22,-3,33,-5,21,0};

            //var result = ClosestToZero(ts);
            //Console.WriteLine(result);
            //Sort(ts);
            */
        private static void DoSomeOtherWork()
        {
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("DoSomeOtherWork");
                Thread.Sleep(500);
            }
        }

        private static void DoSomeWork()
        {
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("DoSomeWork");
                Thread.Sleep(500);
            }
        }

        public static void Fibonacci_Recursive(int len)
        {
            Fibonacci_Rec_Temp(0, 1, 1, len);
        }

        private static void Fibonacci_Rec_Temp(int a, int b, int counter, int len)
        {
            if (counter <= len)
            {
                Console.Write("{0} ", a);
                Fibonacci_Rec_Temp(b, a + b, counter + 1, len);
            }
        }

        public static void Sort(double[] ts)
        {
            for (int i = 0; i < ts.Count(); i++)
            {
                for (int j = i; j < ts.Count(); j++)
                {
                    if (ts[i] > ts[j])
                    {
                        var v = ts[i];
                        ts[i] = ts[j];
                        ts[j] = v;
                    }
                }
            }
        }

        public static double ClosestToZero(double[] ts)
        {
            if (ts.Count() == 0) return 0;
            var r = Math.Abs(ts[0]);
            var val = ts[0];
            for (int i = 1; i < ts.Count(); i++)
            {
                if (Math.Abs(ts[i] - 0) <= r)
                {
                    if (Math.Abs(ts[i]) == val && ts[i]<0)
                    {
                        continue;
                    }
                    r = Math.Abs(ts[i]);
                    val = ts[i];
                }
            }
            return val;
        }

        public static void Pyramid(int level)
        {
            var list = new List<string>();
            var before = new StringBuilder();
            var after = new StringBuilder();
            var conc = new StringBuilder();

            for (int i = 1; i <= level; i++)
            {
                for (int j = level-i; j>0; j--)
                {
                    before.Append(" ");
                }
                for (int j = 1; j < i; j++)
                {
                    before.Append(j);
                }
                for (int j = i; j>1; j--)
                {
                    after.Append(j-1);
                }

                if (before.Length != 0) conc.Append(before.ToString());
                conc.Append(i.ToString());
                if (after.Length != 0) conc.Append(after.ToString());

                list.Add(conc.ToString());
                before.Clear();
                after.Clear();
                conc.Clear();
            }
            foreach (var s in list)
                Console.WriteLine(s);
        }

        public interface icont
        {
            int x { get; }
            int y { set; }

            void faireklk();
        }

        public abstract class abst
        {
            public const int yt = 5;
            public static int varr { get; set; }
            int totla;
            static string tyu { get; }
            protected int ccc;
            protected abstract bool ttt { get; set; }
            public abstract bool bhkg { get; set; }
            abstract public int rest();
            public abstract void rien();
            public int ttttt()
            {
                return 12;
            }
            protected void test()
            {
                Console.WriteLine("test");
            }
        }
        interface IMethods
        {
            void F();
            void G();
        }
        public abstract class C : IMethods
        {
            void IMethods.F() { FF(); }
            void IMethods.G() { GG(); }
            protected abstract void FF();
            protected abstract void GG();
        }
        public class tet2 : C
        {
            public int val;
            protected override void FF()
            {
                throw new NotImplementedException();
            }

            protected override void GG()
            {
                throw new NotImplementedException();
            }
        }
        public class City
        {
            public string name;
            public bool Visited;
            public City(string n)
            {
                name = n;
                Visited = false;
                habitant = new Lazy<string>(GetHab);
            }

            public Lazy<string> habitant;

            private string GetHab()
            {
                Thread.Sleep(4000);
                return "5265";
            }
        }
        public class impAbs : abst, icont
        {
            public readonly List<string> read = new List<string>();
            public readonly tet2 inst = new tet2();
            public const string inst2 = "cfevre";

            public override bool bhkg   // overriding property
            {
                get
                {
                    return false;
                }
                set { }
            }

            public int x
            {
                get
                {
                    return ccc;
                }
            }

            public int y
            {
                set
                {
                    throw new NotImplementedException();
                }
            }

            protected override bool ttt
            {
                get
                {
                    throw new NotImplementedException();
                }

                set
                {
                    throw new NotImplementedException();
                }
            }

            public void faireklk()
            {


                read.Add("ferzget");
                read.AddRange(new List<string> { "12", "13" });
                inst.val = 15;
                inst.val = 16;
                var lst = new List<int> { 0, 1, 2, 3, 4, 5, 6 };
                var res = lst.Where(x => x <= 3);
                lst[1]++;
                foreach (var t in res)
                {
                    Console.WriteLine(t.ToString());
                }
                var cit = new City("Casa");
                Console.WriteLine(cit.habitant.Value);
                var ville = "London";
                IEnumerable<string> names = new[] { "Paris", "Berlin", "London" }; // enumerator
                IEnumerable<City> cities = names.Where(x => x == ville).Select(name => new City(name)); // generator
                ville = "Paris";
                foreach (var city in cities) city.Visited = true;

                foreach (var city in cities)
                {
                    Console.WriteLine(city.Visited + city.name); // prints false
                }
            }

            public sealed override int rest()
            {
                return 1;
            }

            public override void rien()
            {
                throw new NotImplementedException();
            }

            public void testTab(int[] tab, int i = 0)
            {
                tab[0]++;
            }

            public void testTab(string[] tab, int i = 0)
            {
                tab[0] = "val";
            }

            public void testTab(char[] tab, int i = 0)
            {
                tab[0] = 'P';
            }

            public void testTab(string tab, int i)
            {
                tab = "value";
            }

            public void testTab(int i, string tab)
            {
                tab = "value";
            }


        }

        public class cont : icont
        {
            int xx;
            int yy;
            public int x
            {
                get
                {
                    return xx;
                }
            }

            public int y
            {
                set
                {
                    throw new NotImplementedException();
                }
            }

            public void faireklk()
            {
                throw new NotImplementedException();
            }
        }
        #endregion

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message, "Unhandled Thread Exception");
            // here you can log the exception ...
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception).Message, "Unhandled UI Exception");
            // here you can log the exception ...
        }
    }
}
