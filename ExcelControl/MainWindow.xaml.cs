using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.IO;
using System.Collections;
using System.Linq;

namespace ExcelControl
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : Window
	{
		String nameStr = "杜韦洋，萧仪坚，杨右易，陈智惠，黄文泰，刘俊伦，郁珮瑜，胡金瑜，张轩其，陈清灿，傅志鸿，林于庭，金群惠，张俊嘉，吴盈豪，郑盈妍，黄玮谦，李峻湖，于雅萍，王智凡，赵惠萍，童卉清伊柏宇蔡益瑞叶仁尧林郁翔杨梦志黄苑光王柏正黄轩名陈智杰刘佩桦简绍昀邱孟嘉陈孟轩仪张瑞恬吴静如张文华张琼竹卢美玲邱俐音谢佳慧童乐美谢苑佳陈宗翰郭淑慧陈伟伦林育弘黄淑媛蔡琬婷李秀玲郭欣瑜陈俞玄陈志铭钟百军叶倍发表，李仁豪，陈依洁，陈任城，陈静伶，洪海菱，吴怡洁，高文伶，马崇淳，陈冠劭，蔡志铭，刘薇恭，吴淑慧，陈肇旭，谢承颖，邱俊泉，杨雅坚，黄冠郁，黄琼彬，王珈仲，叶静宜，陈雅慧，朱雅铃，王玮鑫，梁幼琦，李柔安，邱冠宇，林佩珊，陈尚仰，林思洁，陈寅恪陈婉恬胡志维郭振芃陈彦姝白彦君谢珮瑜陈雅达陈容辛林雅雯温圣婷郭靖霖蔡淳彬郑承恩王靖邦陈翊伟林怡倩林嘉坚叶玮卿张玮琦游幸琇蔡云辛李丽美陈武江陈奕迅陈奕迅陈奕迅陈奕迅陈奕迅陈绮雯陈钰娴张佑治，徐哲宇，陈静湖，黄智康，蔡伟勇，张介亨，杨致刚，曹佳颖，陈慧湖，吴庆夫，孙俊吉，林郁全，黄书正，徐京容，黄国发，蔡佳其，黄玮秋，钟佳颖，吴治云，元志辛，黄先书，翟伟苹，林政法，林惠芝，谢玉婷，林爱妮，郑婷婷，吴俊贤，梁羽中，林冠富，张采洁，杨怡伶，张晋馨，郑孟儒，陈晓秀，宣佳蓉，谢俊达，张英宸，靳孝皓，吴怡君，黄珍谕，陈威伸，李婉瑜，杜芃均，黎可俊，谢俊德，郑雅圣，许智云，刘正均，陈佳弘，苏冠轩，王珊林，潘乔琪，陈心怡，胡绿倩，陈惟泉，许慈文，叶思妤，林子妹，陈文雄，谢明源，黄孟勋，张瑞麟，林有芳，温凤鑫，张骏易，林育甫，奚宁韦，蒋士铭，萧婉婷，黄思瑜，叶宜静，黄展齐，王丞玟，詹雅辛，张明齐，吴佩璇，许威友，许彦安，蔡建志，陈昱轩，张妍木，彭枝纶，黄婉瑜，林治贵，柯竣裕，林建纬，谢佩旭，林炳雄，连志伟，张雅婷，王雅伯，刘柏恩，王晋琦，林丽君，梁婷婷，韩仰季，朱琬婷，方翰意，蔡万乔，关正斌，张家达，王淑卿，赵胜妹，童纯汉，左彦廷，吴明杰，林素夫，骆钰火，黄郁亨，李倩升，吴健如，杨朝璇，陈美慧，陈欣忠，黄钰伦，连美华，孙以友，赵莹彬，许雅竹，陈明慧，李韦桓，郑丽珠，张健豪，李凯钧，曹建文，李明宏，陈国贵，赖智文，黄夙昀，黄吉玟黎桂盈黄佳玲陈容伟孙佩芬陈淑吉王俊安潘美华徐志明蔡思贤张秀琴李秀琴吴志忠吴梅忠陈盈利宋静怡林晏玮翁隆为杜雅雯陈纹郁黄千喜林怡宜，林惠霖，张淑萍，黄杰，陈佳颖，陈禹洁，刘姵勋，关琼季，王初儒，陈志宜，陈玫洋，刘登阳，阮向靖，毛祥靖，陈文蓁，林慈睿，梁芝妹，李凯元，陈姿妤，金瑜璇，陈金顺，蔡雅琪，傅昶雅，王映恒，吴念任，陈欣怡，刘育如，魏佩芸，黄以佐，徐婷婷，陈宛玲，茹仁豪，黄怡文，林竹娥，陈钰伟，李淑敏，郑钧睿，钟艾桓，陈富意，颜怡君，陈耿裕，黄莉雅，刘佳友，黄丹隆，彭珈文，许怡婷，苏家宏，詹建玮，林尧伦，谢佩辉，蔡登函，蔡婉函，蔡宜良，蔡宜良，吴佩芬，林宗珍，吴奇火，宋协虹，梁淑华，邢信馨，蔡欣毓，林玮妤，潘一东，陈淑彦，吴映桦，林竣纬，吴欣仪，李宜雪，陈金雯，李辰仪，萧家欣，谢明政，封东娟，陈紫宁，林钰和，吴秉生，程宇奇，王丽娟，杨雅琳，许惠瑄，叶恩杰，郭白雄，林绮伶，涂函齐，邱昌达，张俊谚，童佳慧，李绮慈，陈筱谦，邓晓茜，郑志强，陈丽士，刘惠玲，曾佑，方芷杰，陈俐财，柳佳宏，白欣怡，陈雅茜，黄映甫，陈弘旺，李淑玲，李政强，吴涵如，孙旺东，梁夙松，沉昭星，陈凤韵，曹裕仁，高怡婷，张玄屏，张惠瑜，吴淑名，王淑萍，郑文甫，姜婉如，林盈君，李芸瑞，陈雅琪，王俊以，王舒君，钟健豪，姜沛新，郑得臻，陈正昌，陈世容，邓雅婷，郑妍俐，孙治兴，张雅瑜，金俊颖，柯伶骏，袁怡静，刘智伟，杜文豪，范思贤，曾庆圣，姜泰幸，溥民伶，吴秀仪，张伟婷，江怡如，林柏平，张绮雨，蔡夙燕，吴家其，詹贤玮，翁玮婷，黄淑顺，傅佳宁，李宜静，冯俊荣，童欣宜，黄于强，张嘉亨，刘美玲，张秀玲，李淑惠，蔡睿骏，崔凡靖，陈政杰，林圣泉，张美茹，王仕睿，许世柔，赵祥菁，翁佩紫，刘文凯，连采梦，嵇宜臻，陈彦妏，张姿颖，王柏人，张冠航，黄欣怡，张慧婷，蔡靖苹，黎秀英，陈瑞孝，刘士贤，程珊欣，尹其圣，丁淑芬，郭思涵，张江嘉，蒋香奇，陈皇茹，彭志文，洪嘉星，黄琼甫，郑伊迪，杨雅晴，张静如，洪沛琪，许至威，林纯伯，王香君，魏江乔，吕人豪，蓝丽华，黄昱松，林乃志，张伯忠，王诗绮，游慈君，柯建吉，苏世昌，黄志远，郭怡雯，韩佳宏，戴靖儒，江志杰，温俊谚，韩思颖，明哲维，陈慧音，陆亭华，林依男，蔡茜勋，刘萱兰，张雅娟，连千民，李淑芳，韩淑娟，洪丽萍，陈建纬，张哲玮，蔡淑娟，黄嘉虹，林怡璇，朱富宸，张佳桦，林玉珍，赖松山，徐明珠，赖瑶均，刘念雄，曾于婷，萧静雯，王俊美，常思山，白意刚，杨其学，吴珮昌，张嘉芷，魏玮玲，陈仲妹，赵韵安，程博伸，陈志豪，陈丽谕，陈桂木，谢淑玲，刘孟哲，黄淑如，沉世昌，吴嘉盈，李姿城，简宗颖，陈建源，谢欣麟，唐善白，林孟芸，许柏成，侯玉芳，陈莹义，吴文苹，林心怡，王湖均，胡昆玮，方政廷，陈思洁，林虹旭，陈桓仁，李仕邦，陈千紫，何旻慧，黄美珍，唐秀芬，彭恩妃，黄玉婷，黄家合，邓郁雯，王祖善，王俊吉，夏顺芬，袁玉萍，金静宜，周志筠，黄庆蕙，陈韦祥，赖尚盈，张芃薇，蔡明白，王佳仪，邓舒涵童松迪卢伯恬吴美珠赖美玲陈奇廷林婷群苏慧娟陈怡婷陈正芳丁财宣王俞蓁蔡淑贞丁慧娟郑名苓张菁礼杨至明张永发柯勋文李怡君陈彦苓张顺新陈嘉隆，江雅筑，林佳义，吴佳辉，林俊南，林彦霖，林建义，杨雅婷，詹兆萍，冯建宇，吴珮真，虞士玮，赖星胤，李育妮，黄治臻，谢宗诚，李怡伶，黄淑斌，李怡璇，蔡欣桦，谢秀芬，杨诚梅，林佳吟，黄盈君，黄子君，温怡臻，黄慧映，张琼文，陈奕君，陈则伦，王志义，吴韵如，孙金仲，陈信云，陈柏福，许奕恭，杨玫郁，简薇宪，黄幼宁，陈玉谕，杨冠志，张琬欣，林仪麟，周阳娇，朱俊杰，朱宗翰，李良毓，林王汝，杨思颖詹家荣狄启桦林俊政林雅婷王伟妹翁伟杰祖德谷陈仲强廖彦志陈瑞苓罗淑真廖怡伶王丽美林逸明黄圣苹林淳星彭仕慧孙以方杨雯雯杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳杨丞琳林怡萱，陈建强，萧秋萍，黄真妮，王昱义，胡宏伟，苏凯谷，谢宏政，郑怡婷，史敏坚，邱俊良，陈友依，王宥富，邓宜洁，吴诗婷，王雅芬，林耿，，陈佳玲，郑维伦，陈志纯，张甫甫，曾雯雯，陈婉如，周怡君，许力虹，张凯钧，广玮婷，沉玉娟，梁雅婷，胡珈裕，杨奕翔，郑雅惠，林雅婷，林俊冰，谢阳仁，刘孝函，蔡郁婷，金建宏，张嘉昀，周佳蓉，许玮菁，王志明，方俞勇，平佳儒，蓝婉菁，连淑琴，杨倩弘，赖品卉，吴宝恒，陈宗宪，朱培伦，李松柏，施佳静，施婉玲，刘芳妍，陈怡芷，杨文云，蔡依发，蔡孟星，魏智超，洪淑桦，陈玮坚，许雅婷，萧介齐，李美海，陈郁映，许昱宏，张威琦，林心怡，赵佐仲，黄珮欣，王宗翰​​，陈雅雯，吴凯文，林顺定，何玉华，徐怡菁，林依龙，洪雅萍，宋群典，蔡纬谦，梁佳蓉，谢沛惟，李佩霞，陈彦安，蔡紫海，陈凯伦，简惠文，刘怡君，连婷斌，张沛福，杨骏惠，詹慧佐，黄惠如，袁长柔，皮惠如，陈盈天，傅彦霖，谢怡婷，林宜侑，张新颖，马淑真，许柏翰，陈力康，张轩豪，连玄妹，刘美玲，刘大钧，郭佐妹，陈薇爱，蔡佳琪，杜展淳，钱典水，谢美慧，张永玲，吴雅婷，卢惟蕙，林怡昀，林伟伦，连佩桦，林伦群，林佳玟，李世豪，张雅惠，萧易纬，张雅慧，林进其，林雅芳，张向然，吴雅琪，蔡正伟，杨益齐，赖品儒，柳秉勋，林华琦，张禹孝，郑延以，白佳蓉，张淑芸，林育然，蔡雨蕙，余雅婷，王伯宜，宋真妏，廖季一，林宗颖，余惠婷，潘可蓁，陈美惠，张良琳，卢逸湖，黄雯婷，韩景州，吴玉梅，李政秋，黄启城，周雪毓，昱阳，陈昱天，郭家婷，谢伟婷，郑淑雅，陈婷婷，阙俊豪，郭静怡，吴惠珍，柯建宏，陈虹绍，曹世昌，陈佳燕，陈俊维，吴蕙育，吴姿颖，吕正，别雯柔，杜梅恩，蓝彦宏，谢晓妍，周家铭，李欣义，陈奕乔，陈冠良，林松倩，陈柔兴，林家豪，王淑霞，郑宇治，李惠珊，武美玲，陈振琦，陈康铭，吴惠其，赖毓昌，阮孟颖，施佩珊，张心怡，谢宜生，王亭哲，王彦璋，黄家贵，曹建铭，张瑞亦，黄俊强，王家宏，陈志博，林瑞财，张玉玲，林怡君，吴淑君，张永宇，叶佩璇，吴钰雯，陈家恬，林佳杰，萧雅绍，侯明杰，江乃文，黄丽来，许宏儒，吴真亦，郑友茜，张富桓，刘佳彬，吴佩芳，詹俊吉，连丞靖，姜宝轩，温芳仪，王彦儒，吕佳芳，张信豪，戴忠翰，黄威廷，连雅雯，许家凯，林瑞其，黄敏苓，刘品亦，叶琼伟，林玮容，钱国荣，张昕博，黄琬祥，黄惠雯，苏亭君，杨志瑞，邓雅茹，詹圣添，吴良，陈秋新，刘婉幸，陈怡伶，何凯婷，郭小皓，蔡依修，杨秉勋，张丽华，张淑芬，陈映治，钱维伦，潘智杰，沉淑娟，昌莹宏，李松维，赖凡英，黄惠茹，李德宏，许雅雨，骆雅雯，刘佳仪，丁子扬，袁健清，洪慧玲，郭丽萍，何培伦，袁纹仲，陈亚臻，林世彦，陈真娥，吴姿杰，赵君豪，潘胤伦，邓怡君，陈晓瑄，陈佩强，刘千惠，白茂毓，胡如玉，张绮秋，杨孝方，谢俊翔，林孝年，刘怡君，赖喜达，周圣青，刘淑玲，王方铭，林婉瑜，李柏苹，吴丽雯，黄韵升，童俊典，傅雪杰，王思宏，孙凡亨，陈宇轩，黄杰欣，赵季勋，涂建彰，黄正琴，宗淑婷，陈美玲，张宜洁，陈建志，衡乔夫，王至奇，谢建宏，黄子杰，杨士玮，黄清芃，刘文恩，林玮琇，叶怡辉，黄雅婷，杨胤菱，杨家伸，施佩珊，黄怡伶，郭杰馨，张芳明，韩志维，蔡佳慧，戈如如，彭威诚，廖伟喜，黄宜静，邱昱珠，卢子芸";
		public MainWindow()
		{
			InitializeComponent();
			return;
			String[] name = nameStr.Split('，');
			Random r=new Random();
			FileStream fs = new FileStream("ly.csv",FileMode.OpenOrCreate);
			String wxPre = "ot1Ubu";
			int len = 22;
			int Sum = 1500;
			Queue<double> q= GetAS(Sum,367);
			for (int i=0;i<367;i++)
			{//6.30 7.56
				var wxid = wxPre+createRandomChar(r);
				String cost = String.Format("{0:F}",q.Dequeue());
				
				int p = 30+r.Next(60 + 26);
				int h = 16+p / 60;
				int m = p % 60;
				String ts = String.Format("{0}:{1}", h,m.ToString().PadLeft(2, '0'));
				String time = @"2018/2/16 "+ts;
				String s = ",," + wxid + "," + cost + "," + time + Environment.NewLine;
				//Console.WriteLine(ts);
				//Console.WriteLine(cost);
				//Console.WriteLine(wxid);
				Byte[] b = System.Text.Encoding.Default.GetBytes(s);
				fs.Write(b, 0, b.Length);
			}
			
			//fs.Close();
		}
		Queue<double> GetAS(int sum, int siz)
		{
			Random r = new Random();
			Queue<int> q = new Queue<int>();
			long s = 0;
			for (int i = 0; i < siz; i++)
			{
				int k = 1024;
				
				int p = r.Next(3);
				for (int j = 0; j < p; j++)
					k *= 2;
				int num = r.Next(k);
				Console.WriteLine(num);
				s += num;
				q.Enqueue(num);
			}
			//Console.WriteLine(s);
			Queue<double> rec = new Queue<double>();
			for (int i = 0; i < siz; i++)
			{
				double num = q.Dequeue();
				num *= sum;
				num /= s;
				rec.Enqueue(num);
				//Console.WriteLine(num);
			}
			return rec;
		}
		String createRandomChar(Random r)
		{
			StringBuilder sb = new StringBuilder();
			ArrayList c = new ArrayList();
			for (int i = 0; i < 26; i++)
			{
				c.Add((char)('A' + i));
				c.Add((char)('a' + i));
				if (i < 10)
					c.Add((char)('0' + i));
			}
			c.Add('-');
			c.Add('_');
			for (int i = 0; i < 22; i++)
				sb.Append((char)c[r.Next(c.Count)]);
			return sb.ToString();
		}
		private void alert(String s)
		{
			MessageBox.Show(s);
		}
		private void Generate(object sender, RoutedEventArgs e)
		{
			String fileName;
			int recordSum;
			String idPre;
			double minValue;
			double maxValue;
			double sumValue;
			int distributionType;
			DateTime startTime;
			long lastTime;
			if (FileNameBox.Text.Length == 0)
			{
				alert("文件名不能为空");
				return;
			}
			fileName = FileNameBox.Text;
			idPre = IDPreBox.Text;
			try
			{
				recordSum = Int32.Parse(RecordSumBox.Text);
				if (MinValueBox.Text.Length == 0)
					minValue = 0;
				else
					minValue = Double.Parse(MinValueBox.Text);
				if (MaxValueBox.Text.Length == 0)
					maxValue = Int32.MaxValue;
				else
					maxValue= Double.Parse(MaxValueBox.Text);
				if (SumValueBox.Text.Length == 0)
				{
					sumValue = -1;
				}
				else
				{
					sumValue = Double.Parse(SumValueBox.Text);
					maxValue = Math.Max(sumValue / recordSum, maxValue);
					minValue= Math.Min(sumValue / recordSum, minValue);
					maxValue = Math.Min(sumValue, maxValue);
					minValue = Math.Min(sumValue, minValue);
				}
				distributionType = DistributionType.SelectedIndex;
				String[] stime = StartTimeBox.Text.Split(',');
				String[] etime = EndTimeBox.Text.Split(',');
				if (stime.Count() != 5 || etime.Count() != 5)
				{
					alert("时间格式不正确");
					return;
				}
				int[] sArray = new int[5];
				int[] eArray = new int[5];
				for(int i=0;i<5;i++)
				{
					sArray[i] = Int32.Parse(stime[i]);
					eArray[i] = Int32.Parse(etime[i]);
				}
				DateTime sd = new DateTime(sArray[0], sArray[1], sArray[2], sArray[3], sArray[4],0);
				DateTime ed = new DateTime(eArray[0], eArray[1], eArray[2], eArray[3], eArray[4],0);
				startTime = sd;
				lastTime = (int)ed.Subtract(sd).TotalSeconds;
			}
			catch(Exception exception)
			{
				alert("数字不合法，检查是否在数字输入框输入了其他字符");
				return;
			}
			Random r = new Random();
			double[] money = GenerateMoney(minValue, maxValue, sumValue,recordSum, distributionType);
			FileStream fs;
			try
			{
				var fileMode=FileMode.Append;
				switch(FileModeBox.SelectedIndex)
				{
					case 0:
						fileMode = FileMode.Append;
						break;
					case 1:
						fileMode = FileMode.Create;
						break;
				}
				fs = new FileStream(fileName + ".csv",fileMode );
			}catch(IOException exception)
			{
				alert("文件被占用");
				return;
			}
			
			for (int i=0;i<recordSum;i++)
			{
				var wxid = idPre + createRandomChar(r);
				var cost = String.Format("{0:F}", money[i]);
				DateTime randomNowTime = startTime.AddSeconds(r.NextDouble() * lastTime);
				var time = String.Format("{0}/{1}/{2} {3}:{4}:{5}", randomNowTime.Year, randomNowTime.Month, randomNowTime.Day, randomNowTime.Hour, randomNowTime.Minute,randomNowTime.Second);
				String info= String.Format("{0},{1},{2}", wxid, cost, time)+ Environment.NewLine;
				Byte[] b = System.Text.Encoding.Default.GetBytes(info);
				fs.Write(b, 0, b.Length);
			}
			fs.Close();
			alert("生成完成\n见程序目录的文件"+fileName+".csv");
		}
		double[] GenerateMoney(double minValue, double maxValue, double sumValue,int sum,int type)
		{
			double[] array = new double[sum];
			Random r = new Random();
			if (sum == -1)
			{
				for (int i = 0; i < sum; i++)
					array[i] = r.NextDouble() * (maxValue - minValue) + minValue;
				return array;
			}
			double TestSum;
			double VarRate = 0;
			do
			{
				VarRate += 0.05;
				double p = sumValue / sum;
				TestSum = 0;
				for (int i = 0; i < sum; i++)
				{
					double temp = r.NextDouble();
					if (i * 1.0 / sum <= (p - minValue) / (maxValue - minValue))
					{
						array[i] = r.NextDouble() * (maxValue - p) + p;
					}
					else
					{
						array[i] = r.NextDouble() * (p - minValue) + minValue;
					}
					TestSum += array[i];
				}
			}
			while (Math.Abs(TestSum - sumValue)>=sumValue*VarRate);
			//alert(TestSum.ToString());
			double[] recArray = new double[sum];
			bool[] visit = new bool[sum];
			for (int i = 0; i < sum; i++) visit[i] = false;
			for(int i=0;i<sum;i++)
			{
				int index = r.Next(sum - i);
				int cnt = 0;
				for(int j=0;j<sum;j++)
				{
					if (visit[j]) continue;
					if(cnt==index)
					{
						recArray[i] = array[j];
						visit[j] = true;
						break;
					}
					cnt++;
				}
			}
			return recArray;
		}
	}
}
