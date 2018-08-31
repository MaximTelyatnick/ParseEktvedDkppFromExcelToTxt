using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetGear;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;

namespace ParserExcelUKTVEDtoBunariList
{
    public class Program
    {

        static void Main(string[] args)
        {
            string parserType;
            List<ExcelUKTVEDDTO> model = new List<ExcelUKTVEDDTO>();
            List<ExcelDkppDTO> modelDkpp = new List<ExcelDkppDTO>();

            do{

                viewHeader();

                switch (Console.ReadKey().Key)
                {
                    case ConsoleKey.D1:
                        Console.WriteLine("\n");
                        Console.WriteLine("===============================================================================");
                        Console.WriteLine("| Сохранить в файл UKTVED.txt:                                                |");
                        Console.WriteLine("| 1 - Да                                                                      |");
                        Console.WriteLine("| 2 - Нет                                                                     |");
                        Console.WriteLine("| ESC - Отменить                                                              |");
                        Console.WriteLine("===============================================================================");
                        Console.Write("| Сохранить:");
                        switch (Console.ReadKey().Key)
                        {
                            case ConsoleKey.D1:
                                StartParseUktzed(model, true);

                                break;
                            case ConsoleKey.D2:
                                StartParseUktzed(model, false);

                                break;
                            case ConsoleKey.Escape:
                                break;
                        }

                        break;

                    case ConsoleKey.D2:
                        Console.WriteLine("\n");
                        Console.WriteLine("===============================================================================");
                        Console.WriteLine("| Сохранить в файл DKPPP.txt:                                                 |");
                        Console.WriteLine("| 1 - Да                                                                      |");
                        Console.WriteLine("| 2 - Нет                                                                     |");
                        Console.WriteLine("| ESC - Отменить                                                              |");
                        Console.WriteLine("===============================================================================");
                        Console.Write("| Сохранить:");
                        switch (Console.ReadKey().Key)
                        {
                            case ConsoleKey.D1:
                                StartParseDkpp(modelDkpp, true);

                                break;
                            case ConsoleKey.D2:
                                StartParseDkpp(modelDkpp, false);

                                break;
                            case ConsoleKey.Escape:
                                break;
                        }

                        break;



                        break;

                    case ConsoleKey.Escape:
                        Environment.Exit(0);
                        break;

                    default:
                        break;
            }
            }while(true);
 
        }

        static void viewHeader()
        {
            Console.WriteLine("===============================================================================");
            Console.WriteLine("|............___..........______.......__.........__........________..........|");
            Console.WriteLine("|...........//.||.......//......\\......||.........||......||........\\\\........|");
            Console.WriteLine("|..........//..||.......||.............||.........||......||.........||.......|");
            Console.WriteLine("|.........//...||.......\\\\.............||.........||......||_________//.......|");
            Console.WriteLine("|........//....||.........=======......||.........||......||..................|");
            Console.WriteLine("|.......//=====||................\\\\....||.........||......||..................|");
            Console.WriteLine("|......//......||................||....||.........||......||..................|");
            Console.WriteLine("|.....//.......||.......\\\\.......//....||.........||......||..................|");
            Console.WriteLine("|....//........||........=========.....\\\\_________||......||..................|");
            Console.WriteLine("===============================================================================");
            Console.WriteLine("|....................________________________________.........................|");
            Console.WriteLine("|......./\\........../      /\\  ТЕХВАГОНМАШ  /\\       \\........................|");
            Console.WriteLine("|....../__\\......../______/__\\_____________/__\\_______\\.......................|");
            Console.WriteLine("|....../..\\.......|...|_|.......|_|.....|_|.......|_|.|.......................|");
            Console.WriteLine("|...../_.._\\......|...._........._......._........._..|.......................|");
            Console.WriteLine("....../....\\......|...|_|.......|_|.....|_|.......|_|.|.......................|");
            Console.WriteLine("|..../......\\.....|...._........._.. _..._........._..|.......................|");
            Console.WriteLine("|.../___..___\\....|...|_|.......|_|.| |.|_|.......|_|.|.......................|");
            Console.WriteLine("|.......||........|.................| |...............|.......................|");
            Console.WriteLine("===============================================================================");
            Console.WriteLine("| Выберите тип парсера(введите цыфру):                                        |");
            Console.WriteLine("| 1 - ExcelUKTVED                                                             |");
            Console.WriteLine("| 2 - ExcelDKPP                                                               |");
            Console.WriteLine("| ESC - Выйти                                                                 |");
            Console.WriteLine("===============================================================================");
            Console.Write("| Тип парсера: ");
        }

        static int StrAnalizeUktved(String strokaFromExcel)
        {
            int n = 0;

            for (int j = 0; j < strokaFromExcel.Length; j++)
            {
                if (strokaFromExcel[j] != '"' && strokaFromExcel[j] != ' ' && strokaFromExcel[j] != '-')
                {
                    n = j;
                    break;
                }
            }
             n = new Regex("-").Matches(strokaFromExcel.Substring(0, n)).Count;

            return n;
        }

        static void WriteToTxtUktved(List<ExcelUKTVEDDTO> model)
        {
            try
            {
                int i = 0;
                StreamWriter sw = new StreamWriter("UKTVED.txt");

                foreach (var item in model)
                {

                    if (item.ParentId == 0)
                    {
                        Console.WriteLine(item.Code);
                        Console.ReadKey();

                    }
                    switch (item.Level)
                    {
                        case 0:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 1:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 2:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 3:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 4:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 5:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 6:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 7:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 8:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 9:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        default:
                            {

                            }
                            break;
                    }
                }

                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("/n Ошибочка: " + e.Message);
            }
            finally
            {
                Console.WriteLine("/n Отработали, смотри файл.");
            }
        }

        static void WriteToTxtDkpp(List<ExcelDkppDTO> model)
        {
            try
            {
                int i = 0;
                StreamWriter sw = new StreamWriter("DKPP.txt");

                foreach (var item in model)
                {

                    if (item.ParentId == 0)
                    {
                        Console.WriteLine(item.Code);
                        Console.ReadKey();

                    }
                    switch (item.Level)
                    {
                        case 0:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 1:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 2:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 3:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 4:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 5:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        case 6:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 7:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 8:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;
                        case 9:
                            {
                                sw.WriteLine(item.Id + "~" + item.ParentId + "~" + item.Code + "~" + item.Name + "~" + item.Level);
                            }
                            break;

                        default:
                            {

                            }
                            break;
                    }
                }

                sw.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибочка: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Отработали, смотри файл.");
            }
        }
   
        static int SearchParent(String groupName, List<ExcelUKTVEDDTO> model)
        {

            var serachModel = model.SingleOrDefault(m => m.Code == groupName);

            return ((ExcelUKTVEDDTO)serachModel).Id;
        }

        static bool StrTrueOrFalse(String stroka)
        {
            foreach (char c in stroka)
            {
                if (c >= '0' & c <= '9')
                    return false;
            }

            return true;
        }

        static int StrAnalizeDkpp(String strokaFromExcel)
        {
            int n = 0;

            int k = 0;

            if(StrTrueOrFalse(strokaFromExcel))
                return n;

            if(strokaFromExcel.Count()<4)
                return n+1;

            k = strokaFromExcel.Count();

            if (k == 5)
                return n + 2;
            if (k == 6)
                return n + 3;
            if (k == 8)
                return n + 4;
            if (k == 9)
                return n + 5;
            else
                return n + 6; 
        }

        static void StartParseDkpp(List<ExcelDkppDTO> modelDkpp, bool txtWrite)
        {

            var Workbook = Factory.GetWorkbook(@"dkpp_016_2010.xls");
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            int lastLevel = 0, currentLevel = 0, j = 10;
            int lastFirstLevelParent = 0, lastSecondLevelParent = 0, lastThreeLevelParent = 0,
                lastFourthLevelParent = 0, lastFivesLevelParent = 0, lastSixLevelParent = 0, lastSevenLevelParent = 0;

            #region Excel Document to list of ExcelModels

            for (int i = 4; i < 11108; i++)
            {

                lastLevel = currentLevel;

                currentLevel = StrAnalizeDkpp(Convert.ToString(Сells["A" + i].Value));         

                    switch (currentLevel)
                    {
                        case 0:
                            {
                                lastFirstLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = null,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;
                        case 1:
                            {
                                lastSecondLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastFirstLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;
                        case 2:
                            {
                                lastThreeLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastSecondLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;

                        case 3:
                            {
                                lastFourthLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastThreeLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;

                        case 4:
                            {
                                lastFivesLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastFourthLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;

                        case 5:
                            {
                                lastSixLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastFivesLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;
                        case 6:
                            {
                                lastSevenLevelParent = j;
                                modelDkpp.Add(new ExcelDkppDTO()
                                {
                                    Id = j,
                                    Code = Convert.ToString(Сells["A" + i].Value),
                                    Level = currentLevel,
                                    ParentId = lastSixLevelParent,
                                    Name = Convert.ToString(Сells["B" + i].Value),
                                    NameUKTV = Convert.ToString(Сells["C" + i].Value)

                                });
                                ++j;
                            }
                            break;

                        default:
                            {

                            }
                            break;
                    }
            }

            #endregion

            #region List of ExcelModels to Txt
            Console.ReadKey();

            if(txtWrite)
                WriteToTxtDkpp(modelDkpp);

            #endregion

            using (DkppTreeList dkppTreeList = new DkppTreeList(modelDkpp))
            {
                if (dkppTreeList.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {


                }
            }

            


        }

        static void StartParseUktzed(List<ExcelUKTVEDDTO> model, bool txtWrite)
        {
            #region Tittle tables 0 level
            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 1,
                ParentId = -1,
                Code = "Розділ I (01-05)",
                Level = 0,
                Name = "Живi тварини; продукти тваринного походження"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 2,
                ParentId = -1,
                Code = "Розділ II (06-14)",
                Level = 0,
                Name = "Продукти рослинного походження"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 3,
                ParentId = -1,
                Code = "Розділ III (15)",
                Level = 0,
                Name = "Жири та олiї тваринного або рослинного походження; продукти їх розщеплення; готовi харчовi жири; воски тваринного або рослинного походження"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 4,
                ParentId = -1,
                Code = "Розділ IV (16-24)",
                Level = 0,
                Name = "Готовi харчовi продукти; алкогольнi та безалкогольнi напої i оцет; тютюн та його замiнники"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 5,
                Code = "Розділ V (25-27)",
                ParentId = -1,
                Level = 0,
                Name = "Мiнеральнi продукти"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 6,
                Code = "Розділ VI (28-38)",
                ParentId = -1,
                Level = 0,
                Name = "Продукцiя хiмiчної та пов'язаних з нею галузей промисловостi"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 7,
                Code = "Розділ VII (39-40)",
                ParentId = -1,
                Level = 0,
                Name = "Полiмернi матерiали, пластмаси та вироби з них; каучук, гума та вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 8,
                Code = "Розділ VIII (41-43)",
                ParentId = -1,
                Level = 0,
                Name = "Шкури необробленi, шкiра вичинена, натуральне та штучне хутро та вироби з них; шорно-сiдельнi вироби та упряж; дорожнi речi, сумки та аналогiчнi товари; вироби з кишок тварин (крiм кетгуту з натурального шовку)"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 9,
                Code = "Розділ IX (44-46)",
                ParentId = -1,
                Level = 0,
                Name = "Деревина i вироби з деревини; деревне вугiлля; корок та вироби з нього; вироби iз соломи, альфи та iнших матерiалiв для плетiння; кошиковi та iншi плетенi вироби"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 10,
                Code = "Розділ X (47-49)",
                ParentId = -1,
                Level = 0,
                Name = "Маса з деревини або з iнших волокнистих целюлозних матерiалiв; папiр або картон, одержанi з вiдходiв та макулатури; папiр, картон та вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 11,
                Code = "Розділ XI (50-63)",
                ParentId = -1,
                Level = 0,
                Name = "Текстильнi матерiали та текстильнi вироби"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 12,
                Code = "Розділ XII (64-67)",
                ParentId = -1,
                Level = 0,
                Name = "Взуття, головнi убори, парасольки вiд дощу та сонця, палицi, стеки, батоги та їх частини; пiр'я оброблене i вироби з нього; штучнi квiти; вироби з волосся людини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 13,
                Code = "Розділ XIII (68-70)",
                ParentId = -1,
                Level = 0,
                Name = "Вироби з каменю, гiпсу, цементу, азбесту, слюди або аналогiчних матерiалiв; керамiчнi вироби; скло та вироби iз скла"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 14,
                Code = "Розділ XIV (71)",
                ParentId = -1,
                Level = 0,
                Name = "Перли природнi або культивованi, дорогоцiнне або напiвдорогоцiнне камiння, дорогоцiннi метали, метали, плакованi дорогоцiнними металами, та вироби з них; бiжутерiя; монети"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 15,
                Code = "Розділ XV (72-83)",
                ParentId = -1,
                Level = 0,
                Name = "Недорогоцiннi метали та вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 16,
                Code = "Розділ XVI (84-85)",
                ParentId = -1,
                Level = 0,
                Name = "Машини, обладнання та механiзми; електротехнiчне обладнання; їх частини; звукозаписувальна та звуковiдтворювальна апаратура, апаратура для запису або вiдтворення телевiзiйного зображення i звуку, їх частини та приладдя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 18,
                Code = "Розділ XVII (86-89)",
                ParentId = -1,
                Level = 0,
                Name = "Засоби наземного транспорту, лiтальнi апарати, плавучi засоби i пов'язанi з транспортом пристрої та обладнання"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 19,
                Code = "Розділ XVIII (90-92)",
                ParentId = -1,
                Level = 0,
                Name = "Прилади та апарати оптичнi, фотографiчнi, кiнематографiчнi, контрольнi, вимiрювальнi, прецизiйнi, медичнi або хiрургiчнi; годинники всiх видiв; музичнi iнструменти; їх частини та приладдя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 20,
                Code = "Розділ XIX (93)",
                ParentId = -1,
                Level = 0,
                Name = "Зброя, боєприпаси; їх частини та приладдя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 21,
                Code = "Розділ XX (94-96)",
                ParentId = -1,
                Level = 0,
                Name = "Рiзнi промисловi товари"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 22,
                Code = "Розділ XXI (97)",
                ParentId = -1,
                Level = 0,
                Name = "Твори мистецтва, предмети колекцiонування та антикварiат"
            });
            #endregion

            #region Розділ I (01-05) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 23,
                Code = "Група 01",
                ParentId = 1,
                Level = 1,
                Name = "Живi тварини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 24,
                Code = "Група 02",
                ParentId = 1,
                Level = 1,
                Name = "М'ясо та їстiвнi субпродукти"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 25,
                Code = "Група 03",
                ParentId = 1,
                Level = 1,
                Name = "Риба i ракоподiбнi, молюски та iншi водянi безхребетнi"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 26,
                Code = "Група 04",
                ParentId = 1,
                Level = 1,
                Name = "Молоко та молочнi продукти; яйця птицi; натуральний мед; їстiвнi продукти тваринного походження, в iншому мiсцi не зазначенi"
            });


            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 27,
                Code = "Група 05",
                ParentId = 1,
                Level = 0,
                Name = "Iншi продукти тваринного походження, в iншому мiсцi не зазначенi"
            });

            #endregion

            #region Розділ II (06-14) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 28,
                Code = "Група 06",
                ParentId = 2,
                Level = 1,
                Name = "Живi дерева та iншi рослини; цибулини, корiння та iншi аналогiчнi частини рослин; зрiзанi квiти i декоративна зелень"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 29,
                Code = "Група 07",
                ParentId = 2,
                Level = 1,
                Name = "Овочi та деякi їстiвнi коренеплоди i бульби"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 30,
                Code = "Група 08",
                ParentId = 2,
                Level = 1,
                Name = "Їстiвнi плоди та горiхи; шкiрки цитрусових або динь"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 31,
                Code = "Група 09",
                ParentId = 2,
                Level = 1,
                Name = "Кава, чай, мате або парагвайський чай, прянощi"
            });


            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 32,
                Code = "Група 10",
                ParentId = 2,
                Level = 1,
                Name = "Зерновi культури"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 33,
                Code = "Група 11",
                ParentId = 2,
                Level = 1,
                Name = "Продукцiя борошномельно-круп'яної промисловостi; солод; крохмалi; iнулiн; пшенична клейковина"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 34,
                Code = "Група 12",
                ParentId = 2,
                Level = 1,
                Name = "Насiння i плоди олiйних рослин; iнше насiння, плоди та зерна; технiчнi або лiкарськi рослини; солома i фураж"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 35,
                Code = "Група 13",
                ParentId = 2,
                Level = 1,
                Name = "Шелак природний неочищений; камедi, смоли та iншi рослиннi соки i екстракти"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 36,
                Code = "Група 14",
                ParentId = 2,
                Level = 1,
                Name = "	Рослиннi матерiали для виготовлення плетених виробiв; iншi продукти рослинного походження, в iншому мiсцi не зазначенi"
            });

            #endregion

            #region Розділ III (15) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 37,
                Code = "Група 15",
                ParentId = 3,
                Level = 1,
                Name = "Жири та олiї тваринного або рослинного походження; продукти їх розщеплення; готовi харчовi жири; воски тваринного або рослинного походження"
            });



            #endregion

            #region Розділ IV (16-24) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 38,
                Code = "Група 16",
                ParentId = 4,
                Level = 1,
                Name = "Готовi харчовi продукти з м'яса, риби або ракоподiбних, молюскiв або iнших водяних безхребетних"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 39,
                Code = "Група 17",
                ParentId = 4,
                Level = 1,
                Name = "Цукор i кондитерськi вироби з цукру"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 40,
                Code = "Група 18",
                ParentId = 4,
                Level = 1,
                Name = "Какао та продукти з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 41,
                Code = "Група 19",
                ParentId = 4,
                Level = 1,
                Name = "Готовi продукти iз зерна зернових культур, борошна, крохмалю або молока; борошнянi кондитерськi вироби"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 42,
                Code = "Група 20",
                ParentId = 4,
                Level = 1,
                Name = "Продукти переробки овочiв, плодiв, горiхiв або iнших частин рослин"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 43,
                Code = "Група 21",
                ParentId = 4,
                Level = 1,
                Name = "Рiзнi харчовi продукти"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 44,
                Code = "Група 22",
                ParentId = 4,
                Level = 1,
                Name = "Алкогольнi i безалкогольнi напої та оцет"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 45,
                Code = "Група 23",
                ParentId = 4,
                Level = 1,
                Name = "Залишки i вiдходи харчової промисловостi; готовi корми для тварин"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 46,
                Code = "Група 24",
                ParentId = 4,
                Level = 1,
                Name = "Тютюн i промисловi замiнники тютюну"
            });



            #endregion

            #region Розділ V (25-27) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 47,
                Code = "Група 25",
                ParentId = 5,
                Level = 1,
                Name = "Сiль; сiрка; землi та камiння; штукатурнi матерiали, вапно та цемент"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 48,
                Code = "Група 26",
                ParentId = 5,
                Level = 1,
                Name = "Руди, шлак i зола"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 49,
                Code = "Група 27",
                ParentId = 5,
                Level = 1,
                Name = "Палива мiнеральнi; нафта i продукти її перегонки; бiтумiнознi речовини; воски мiнеральнi"
            });

            #endregion

            #region Розділ VI (28-38) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 50,
                Code = "Група 28",
                ParentId = 6,
                Level = 1,
                Name = "Продукти неорганiчної хiмiї: неорганiчнi або органiчнi сполуки дорогоцiнних металiв, рiдкiсноземельних металiв, радiоактивних елементiв або iзотопiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 51,
                Code = "Група 29",
                ParentId = 6,
                Level = 1,
                Name = "Органiчнi хiмiчнi сполуки"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 52,
                Code = "Група 30",
                ParentId = 6,
                Level = 1,
                Name = "Фармацевтична продукцiя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 53,
                Code = "Група 31",
                ParentId = 6,
                Level = 1,
                Name = "Добрива"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 54,
                Code = "Група 32",
                ParentId = 6,
                Level = 1,
                Name = "Екстракти дубильнi або барвнi; танiни та їх похiднi, барвники, пiгменти та iншi фарбувальнi матерiали, фарби i лаки; замазки та iншi мастики; чорнило, туш"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 55,
                Code = "Група 33",
                ParentId = 6,
                Level = 1,
                Name = "Ефiрнi олiї та резиноїди; парфумернi, косметичнi та туалетнi препарати"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 56,
                Code = "Група 34",
                ParentId = 6,
                Level = 1,
                Name = "Мило, поверхнево-активнi органiчнi речовини, мийнi засоби, мастильнi матерiали, воски штучнi та готовi, сумiшi для чищення або полiрування, свiчки та аналогiчнi вироби, пасти для лiплення, пластилiн, 'стоматологiчний вiск' i сумiшi на основi гiпсу для стоматологiї"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 57,
                Code = "Група 35",
                ParentId = 6,
                Level = 1,
                Name = "Бiлковi речовини; модифiкованi крохмалi; клеї; ферменти"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 58,
                Code = "Група 36",
                ParentId = 6,
                Level = 1,
                Name = "Порох i вибуховi речовини; пiротехнiчнi вироби; сiрники; пiрофорнi сплави; деякi горючi матерiали"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 59,
                Code = "Група 37",
                ParentId = 6,
                Level = 1,
                Name = "Фотографiчнi або кiнематографiчнi товари"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 60,
                Code = "Група 38",
                ParentId = 6,
                Level = 1,
                Name = "Рiзноманiтна хiмiчна продукцiя"
            });

            #endregion

            #region Розділ VII (39-40) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 61,
                Code = "Група 39",
                ParentId = 7,
                Level = 1,
                Name = "Пластмаси, полiмернi матерiали та вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 62,
                Code = "Група 40",
                ParentId = 7,
                Level = 1,
                Name = "Каучук, гума та вироби з них"
            });

            #endregion

            #region Розділ VIII (41-43) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 63,
                Code = "Група 41",
                ParentId = 8,
                Level = 1,
                Name = "Шкури необробленi (крiм натурального та штучного хутра) i шкiра вичинена"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 64,
                Code = "Група 42",
                ParentId = 8,
                Level = 1,
                Name = "Вироби iз шкiри; шорно-сiдельнi вироби та упряж; дорожнi речi, сумки та аналогiчнi товари; вироби з кишок тварин (крiм кетгуту з натурального шовку)"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 65,
                Code = "Група 43",
                ParentId = 8,
                Level = 1,
                Name = "Натуральне та штучне хутро; вироби з нього"
            });

            #endregion

            #region Розділ IX (44-46) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 66,
                Code = "Група 44",
                ParentId = 9,
                Level = 1,
                Name = "Деревина i вироби з деревини; деревне вугiлля"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 67,
                Code = "Група 45",
                ParentId = 9,
                Level = 1,
                Name = "Корок та вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 68,
                Code = "Група 46",
                ParentId = 9,
                Level = 1,
                Name = "Вироби iз соломи, трави еспарто та iнших матерiалiв, якi використовуються для плетiння; кошиковi вироби та плетенi вироби"
            });

            #endregion

            #region Розділ X (47-49) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 69,
                Code = "Група 47",
                ParentId = 10,
                Level = 1,
                Name = "Маса з деревини або з iнших волокнистих целюлозних матерiалiв; папiр або картон, одержанi з вiдходiв та макулатури"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 70,
                Code = "Група 48",
                ParentId = 10,
                Level = 1,
                Name = "Папiр i картон; вироби з паперової маси, паперу або картону"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 71,
                Code = "Група 49",
                ParentId = 10,
                Level = 1,
                Name = "Друкована продукцiя, перiодичнi видання або iнша продукцiя полiграфiчної промисловостi; рукописи або машинописнi тексти та плани"
            });

            #endregion

            #region Розділ XI (50-63) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 72,
                Code = "Група 50",
                ParentId = 11,
                Level = 1,
                Name = "Шовк"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 73,
                Code = "Група 51",
                ParentId = 11,
                Level = 1,
                Name = "Вовна, тонкий та грубий волос тварин; пряжа i тканини з кiнського волосу"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 74,
                Code = "Група 52",
                ParentId = 11,
                Level = 1,
                Name = "Бавовна"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 75,
                Code = "Група 53",
                ParentId = 11,
                Level = 1,
                Name = "Iншi рослиннi текстильнi волокна; пряжа з паперу i тканини з паперової пряжi"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 76,
                Code = "Група 54",
                ParentId = 11,
                Level = 1,
                Name = "Нитки синтетичнi або штучнi; стрiчковi та подiбної форми нитки iз синтетичних або штучних матерiалiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 77,
                Code = "Група 55",
                ParentId = 11,
                Level = 1,
                Name = "Синтетичнi або штучнi штапельнi волокна"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 78,
                Code = "Група 56",
                ParentId = 11,
                Level = 1,
                Name = "Вата, повсть i нетканi матерiали; спецiальна пряжа; шпагати, мотузки, троси та канати i вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 79,
                Code = "Група 57",
                ParentId = 11,
                Level = 1,
                Name = "Килими та iншi текстильнi покриття для пiдлоги"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 80,
                Code = "Група 58",
                ParentId = 11,
                Level = 1,
                Name = "Спецiальнi тканини; тафтинговi текстильнi матерiали; мережива; гобелени; оздоблювальнi матерiали; вишивка"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 81,
                Code = "Група 59",
                ParentId = 11,
                Level = 1,
                Name = "Текстильнi матерiали, просоченi, покритi або дубльованi; текстильнi вироби технiчного призначення"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 82,
                Code = "Група 60",
                ParentId = 11,
                Level = 1,
                Name = "Трикотажнi полотна"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 83,
                Code = "Група 61",
                ParentId = 11,
                Level = 1,
                Name = "Одяг та додатковi речi до одягу, трикотажнi"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 84,
                Code = "Група 62",
                ParentId = 11,
                Level = 1,
                Name = "Одяг та додатковi речi до одягу, текстильнi, крiм трикотажних"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 85,
                Code = "Група 63",
                ParentId = 11,
                Level = 1,
                Name = "Iншi готовi текстильнi вироби; набори; одяг та текстильнi вироби, що використовувались; ганчiр'я"
            });

            #endregion

            #region Розділ XII (64-67) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 86,

                Code = "Група 64",
                ParentId = 12,
                Level = 1,
                Name = "Взуття, гетри та аналогiчнi вироби; їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 87,
                Code = "Група 65",
                ParentId = 12,
                Level = 1,
                Name = "Головнi убори та їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 88,
                Code = "Група 66",
                ParentId = 12,
                Level = 1,
                Name = "Парасольки вiд дощу та сонця, палицi, палицi-сидiння, батоги, хлисти для верхової їзди та їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 89,
                Code = "Група 67",
                ParentId = 12,
                Level = 1,
                Name = "Обробленi пiр'я та пух i вироби з них; штучнi квiти; вироби з волосся людини"
            });

            #endregion

            #region Розділ XIII (68-70) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 90,
                Code = "Група 68",
                ParentId = 13,
                Level = 1,
                Name = "Вироби з каменю, гiпсу, цементу, азбесту, слюди або аналогiчних матерiалiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 91,
                Code = "Група 69",
                ParentId = 13,
                Level = 1,
                Name = "Керамiчнi вироби"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 92,
                Code = "Група 70",
                ParentId = 13,
                Level = 1,
                Name = "Скло та вироби iз скла"
            });

            #endregion

            #region Розділ XIV (71) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 93,
                Code = "Група 71",
                ParentId = 14,
                Level = 1,
                Name = "Перли природнi або культивованi, дорогоцiнне або напiвдорогоцiнне камiння, дорогоцiннi метали, метали, плакованi дорогоцiнними металами, та вироби з них; бiжутерiя; монети"
            });

            #endregion

            #region Розділ XV (72-83) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 94,
                Code = "Група 72",
                ParentId = 15,
                Level = 1,
                Name = "Чорнi метали"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 95,
                Code = "Група 73",
                ParentId = 15,
                Level = 1,
                Name = "Вироби з чорних металiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 96,
                Code = "Група 74",
                ParentId = 15,
                Level = 1,
                Name = "Мiдь i вироби з неї"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 97,
                Code = "Група 75",
                ParentId = 15,
                Level = 1,
                Name = "Нiкель i вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 98,
                Code = "Група 76",
                ParentId = 15,
                Level = 1,
                Name = "Алюмiнiй i вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 99,
                Code = "Група 78",
                ParentId = 15,
                Level = 1,
                Name = "Свинець i вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 100,
                Code = "Група 79",
                ParentId = 15,
                Level = 1,
                Name = "Цинк i вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 101,
                Code = "Група 80",
                ParentId = 15,
                Level = 1,
                Name = "Олово i вироби з нього"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 102,
                Code = "Група 81",
                ParentId = 15,
                Level = 1,
                Name = "Iншi недорогоцiннi метали; металокерамiка; вироби з них"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 103,
                Code = "Група 82",
                ParentId = 15,
                Level = 1,
                Name = "Iнструменти, ножовi вироби, ложки та виделки з недорогоцiнних металiв; їх частини з недорогоцiнних металiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 104,
                Code = "Група 83",
                ParentId = 15,
                Level = 1,
                Name = "Iншi вироби з недорогоцiнних металiв"
            });

            #endregion

            #region Розділ XVI (84-85) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 105,
                Code = "Група 84",
                ParentId = 16,
                Level = 1,
                Name = "Реактори ядернi, котли, машини, обладнання i механiчнi пристрої; їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 106,
                Code = "Група 85",
                ParentId = 16,
                Level = 1,
                Name = "Електричнi машини, обладнання та їх частини; апаратура для запису або вiдтворення звуку; телевiзiйна апаратура для запису та вiдтворення зображення i звуку, їх частини та приладдя"
            });

            #endregion

            #region Розділ XVII (86-89) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 107,
                Code = "Група 86",
                ParentId = 18,
                Level = 1,
                Name = "Залiзничнi локомотиви або моторнi вагони трамвая, рухомий склад та їх частини; шляхове обладнання та пристрої для залiзниць або трамвайних колiй та їх частини; механiчне (у тому числi електромеханiчне) сигналiзацiйне обладнання всiх видiв"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 108,
                Code = "Група 87",
                ParentId = 18,
                Level = 1,
                Name = "Засоби наземного транспорту, крiм залiзничного або трамвайного рухомого складу, їх частини та обладнання"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 109,
                Code = "Група 88",
                ParentId = 18,
                Level = 1,
                Name = "Лiтальнi апарати, космiчнi апарати та їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 110,
                Code = "Група 89",
                ParentId = 18,
                Level = 1,
                Name = "Судна, човни та iншi плавучi засоби"
            });

            #endregion

            #region Розділ XVIII (90-92) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 111,
                Code = "Група 90",
                ParentId = 19,
                Level = 1,
                Name = "Прилади та апарати оптичнi, фотографiчнi, кiнематографiчнi, контрольнi, вимiрювальнi, прецизiйнi; медичнi або хiрургiчнi; їх частини та приладдя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 112,
                Code = "Група 91",
                ParentId = 19,
                Level = 1,
                Name = "Годинники всiх видiв та їх частини"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 113,
                Code = "Група 92",
                ParentId = 19,
                Level = 1,
                Name = "Музичнi iнструменти; їх частини та приладдя"
            });

            #endregion

            #region Розділ XIX (93) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 114,
                Code = "Група 93",
                ParentId = 20,
                Level = 1,
                Name = "Зброя, боєприпаси; їх частини та приладдя"
            });

            #endregion

            #region Розділ XX (94-96) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 115,
                Code = "Група 94",
                ParentId = 21,
                Level = 1,
                Name = "Меблi; постiльнi речi, матраци, матрацнi основи, диваннi подушки та аналогiчнi набивнi речi меблiв, свiтильники та освiтлювальнi прилади, в iншому мiсцi не зазначенi; свiтловi покажчики, табло та подiбнi вироби; збiрнi будiвельнi конструкцiї"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 116,
                Code = "Група 95",
                ParentId = 20,
                Level = 1,
                Name = "Iграшки, iгри та спортивний iнвентар; їх частини та приладдя"
            });

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 117,
                Code = "Група 96",
                ParentId = 20,
                Level = 1,
                Name = "Рiзнi готовi вироби"
            });

            #endregion

            #region Розділ XXI (97) 1 level

            model.Add(new ExcelUKTVEDDTO()
            {
                Id = 118,
                Code = "Група 97",
                ParentId = 22,
                Level = 1,
                Name = "	Твори мистецтва, предмети колекцiонування та антикварiат"
            });

            #endregion


            var Workbook = Factory.GetWorkbook(@"uktzed_2013.xlsx");
            var Worksheet = Workbook.Worksheets[0];
            var Сells = Worksheet.Cells;
            int lastLevel = 0, currentLevel = 0, j = 119;
            int  lastFirstLevelParent = 0, lastSecondLevelParent = 0, lastThreeLevelParent = 0,
                lastFourthLevelParent = 0, lastFivesLevelParent = 0, lastSixLevelParent = 0, lastSevenLevelParent = 0;

            #region Excel Document to list of ExcelModels

            for (int i = 3; i < 16395; i++)
            {

                lastLevel = currentLevel;

                currentLevel = StrAnalizeUktved(Convert.ToString(Сells["B" + i].Value));

                if (currentLevel - lastLevel > 1)
                {
                    
                    Console.WriteLine(Convert.ToString(Сells["A" + i].Value));
                    Console.WriteLine("|" + currentLevel + "|" + lastLevel + "|");
                    Console.ReadKey();
                }
           
                switch (currentLevel)
                {
                    case 0:
                        {
                            lastFirstLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {   
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel+2,
                                ParentId = SearchParent("Група " + Convert.ToString(Сells["A" + i].Value).Substring(0, 2), model),
                                Name = Convert.ToString(Сells["B" + i].Value)
                            });
                            ++j;
                        }
                        break;
                    case 1:
                        {
                            lastSecondLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel+2,
                                ParentId = lastFirstLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;
                    case 2:
                        {
                            lastThreeLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            { 
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel + 2,
                                ParentId = lastSecondLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;

                    case 3:
                        {
                            lastFourthLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {   
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel + 2,
                                ParentId = lastThreeLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;

                    case 4:
                        {
                            lastFivesLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel + 2,
                                ParentId = lastFourthLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;

                    case 5:
                        {
                            lastSixLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel + 2,
                                ParentId = lastFivesLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;
                    case 6:
                        {
                            lastSevenLevelParent = j;
                            model.Add(new ExcelUKTVEDDTO()
                            {
                                Id = j,
                                Code = Convert.ToString(Сells["A" + i].Value),
                                Level = currentLevel + 2,
                                ParentId = lastSixLevelParent,
                                Name = Convert.ToString(Сells["B" + i].Value)

                            });
                            ++j;
                        }
                        break;

                    default:
                        {

                        }
                        break;
                }
            }

            #endregion

            #region List of ExcelModels to Txt

            if(txtWrite)
                WriteToTxtUktved(model);

            #endregion

            using (UktvedTreeList uktvedTreeList = new UktvedTreeList(model))
            {
                if (uktvedTreeList.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {


                }
            }
            
            Console.ReadKey();   
        }
        
    }
}
