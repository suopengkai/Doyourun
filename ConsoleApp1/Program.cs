using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            TestPlinq();
            var beatles = (new[]
                            {   new { id=1 , inst = "guitar" , name="john" },
                                new { id=3 , inst = "drums" , name="george" },
                                new { id=4 , inst = "guitar" , name="paul" },
                                new { id=2 , inst = "drums" , name="ringo" },
                                new { id=5 , inst = "drums" , name="pete" }
                            }
                        );

            var a = (from b1 in beatles
                     where b1.inst == "guitar"
                     select b1).ToString();

            var w = new Stopwatch();
            w.Start();
            var c = from b in beatles orderby b.id descending group b by b.inst into b select new { c = b.FirstOrDefault().id, a = b.FirstOrDefault().inst, d = b.FirstOrDefault().name, vv = b.First().name };
            foreach (var item in c)
            {
                Console.WriteLine(item.vv);
            }
            var lingtosql = a.ToString();
            w.Stop();
            Console.WriteLine("linqtosql:=/n{0},时间{1}", a, w.ElapsedMilliseconds);

            #region sql----ROW_Number() over (partition by inst order by id)

            var o = beatles.OrderBy(m => m.id).ToArray().GroupBy(x => x.inst)
                .SelectMany(t => t.Select((b, i) => new { b, i })).Select(m => m.b).ToList();
            foreach (var item in o)
            {
                Console.WriteLine(item.id);
            }
            var o1 = beatles.OrderBy(x => x.id).GroupBy(x => x.inst)
                .Select(g => new { g, count = g.Count() }).ToList()
                .SelectMany(t => t.g.Select(b => b).Zip(Enumerable.Range(1, t.count), (j, i) => new { j.inst, j.name, rn = i }));
            foreach (var item in o1)
            {
                Console.WriteLine(item.rn);
            }
            #endregion
            Console.WriteLine("======");
            var number = Getnumber();

            Console.WriteLine(number.Sum().ToString());
            Console.WriteLine(number.Sum().ToString());

            Console.ReadLine();
        }
        static IEnumerable<int> Getnumber()
        {
            //取 0-3 直接的随机数子
            var counr = rand.Next(0, 3);
            return Enumerable.Range(0, counr).Select(n => rand.Next(0, 10));
        }
      public   static Random rand = new Random();
        public static void TestPlinq()
        {
            ///字典  线程安全
            var dic = new ConcurrentDictionary<string, Dog>();
            Parallel.For(0, 1000000, (i) =>
           {
               var dog = new Dog()
               { MyProperty = i, Time = DateTime.Now.AddSeconds(i), Year = new Random().Next(1, 10) };
               dic.TryAdd("狗"+i,dog);
           });
            Console.WriteLine("插入{0}条",dic.Count);

            ///开始查询
            var wath = new Stopwatch();
            wath.Start();
            //asParaller  分析查询总数据结构 通过并行查询可能提高查询速度   plinq 将原序列 分为可以同时执行的认为 如果并行不安全就按照原顺序
            var w = (from n in dic.Values.AsParallel() where n.Year > 3 && n.Year < 7 select n).ToList();
            wath.Stop();
            Console.WriteLine($"用了 plinq {wath.ElapsedMilliseconds}");

            wath.Start();
            var c = from n in dic.Values where n.Year > 3 && n.Year < 7 select n;
            wath.Stop();
            Console.WriteLine($" 普通 linq {wath.ElapsedMilliseconds}");


        }



        public class Dog
        {
            public int MyProperty { get; set; }
            public int Year { get; set; }
            public DateTime Time { get; set; }

        }
    }
}
