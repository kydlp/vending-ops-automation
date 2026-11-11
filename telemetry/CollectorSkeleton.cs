using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;

// 収集側の骨格。
//
// ★ ベンダー依存部は ISalesSource の裏に隔離してある。実装（StubSalesSource）は
//   固定値を返すだけのスタブで、実案件のベンダー製品を叩くコードは含まない。
//   製品固有のクラス名・メソッド呼び出し順・エンドポイント形式は非公開。
//   → docs/telemetry-approach.md
//
// ここに残してあるのは「ベンダーが何であっても要る部分」：
//   設定の読み込み／リトライ／0件を正常系にしない検証／TSV整形／通知／終了コード。
namespace Telemetry
{
    /// <summary>1号機・1日ぶんの売上。</summary>
    internal sealed class SalesRow
    {
        public string Measure;      // 検量区分（夜間 / 回収時）
        public DateTime Date;
        public int MachineId;
        public string MachineName;
        public int Units;
        public int Total;
        public int Cash;
        public int Card;

        public string ToTsv()
        {
            return string.Join("\t", new[]
            {
                Measure,
                Date.ToString("yyyy-MM", CultureInfo.InvariantCulture),
                Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                MachineId.ToString(CultureInfo.InvariantCulture),
                MachineName,
                Units.ToString(CultureInfo.InvariantCulture),
                Total.ToString(CultureInfo.InvariantCulture),
                Cash.ToString(CultureInfo.InvariantCulture),
                Card.ToString(CultureInfo.InvariantCulture),
            });
        }
    }

    /// <summary>
    /// 売上の取得元。ベンダー固有の接続はこの裏に閉じ込める。
    /// 実装を差し替えれば、別システムでも上位の処理はそのまま使える。
    /// </summary>
    internal interface ISalesSource
    {
        void Connect(Config config);
        IList<SalesRow> FetchDaily(DateTime day);
    }

    /// <summary>公開用のスタブ。実接続は含まない。</summary>
    internal sealed class StubSalesSource : ISalesSource
    {
        public void Connect(Config config)
        {
            if (string.IsNullOrEmpty(config.Endpoint))
                throw new InvalidOperationException("endpoint が未設定です。");
            // 実装では、ベンダーの許可を得たうえで公式クライアントのライブラリを使い、
            // 公式クライアントと同じ手順で初期化・ログインする。
            // ログインで得た情報を後続の呼び出しに渡し忘れると「エラーにならず0件」という壊れ方をする。
        }

        public IList<SalesRow> FetchDaily(DateTime day)
        {
            // 結果の有無は実際の件数で判断する（「データあり」の表示だけを信じない）。
            return new List<SalesRow>
            {
                new SalesRow { Measure = "夜間", Date = day, MachineId = 1,
                               MachineName = "Site01", Units = 27, Total = 22950,
                               Cash = 10200, Card = 12750 },
            };
        }
    }

    internal static class Collector
    {
        private const int MaxAttempts = 3;

        private static int Main(string[] args)
        {
            var day = args.Length > 0
                ? DateTime.ParseExact(args[0], "yyyy-MM-dd", CultureInfo.InvariantCulture)
                : DateTime.Today.AddDays(-1);

            Config config;
            try
            {
                config = Config.Load();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("設定の読み込みに失敗: " + e.Message);
                return 1;
            }

            IList<SalesRow> rows;
            try
            {
                rows = FetchWithRetry(new StubSalesSource(), config, day);
            }
            catch (Exception e)
            {
                Notify(config, string.Format("売上取得に失敗しました（{0:yyyy-MM-dd}）: {1}",
                                             day, e.Message));
                Console.Error.WriteLine(e);
                return 1;
            }

            var warnings = Validate(rows);
            WriteTsv(rows, string.Format("sales_{0:yyyyMMdd}.tsv", day));

            var summary = Summarize(day, rows, warnings);
            Notify(config, summary);
            Console.WriteLine(summary);

            // 警告があっても出力はする。黙って止まるより、出したうえで知らせる。
            return 0;
        }

        private static IList<SalesRow> FetchWithRetry(ISalesSource source, Config config, DateTime day)
        {
            Exception last = null;
            for (int attempt = 1; attempt <= MaxAttempts; attempt++)
            {
                try
                {
                    source.Connect(config);
                    return source.FetchDaily(day);
                }
                catch (Exception e)
                {
                    last = e;
                    Console.Error.WriteLine(string.Format("取得失敗 {0}/{1}: {2}",
                                                          attempt, MaxAttempts, e.Message));
                    if (attempt < MaxAttempts)
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(10 * attempt));
                }
            }
            throw new Exception("リトライ上限に達しました。", last);
        }

        /// <summary>
        /// 0件・全台ゼロを正常系として通さない。ここが無いと、
        /// ログイン情報の渡し漏れのような「静かな失敗」が無警告で通過する。
        /// </summary>
        private static IList<string> Validate(IList<SalesRow> rows)
        {
            var warnings = new List<string>();
            if (rows.Count == 0)
            {
                warnings.Add("取得0件。接続は成功しているが中身が空＝ログイン情報の受け渡し不備の疑い");
                return warnings;
            }

            int total = 0;
            foreach (var r in rows)
            {
                total += r.Total;
                if (r.Cash + r.Card != r.Total)
                    warnings.Add(string.Format("号機{0}: 現金+カードが総額と一致しない", r.MachineId));
                if (r.Units > 0 && r.Total == 0)
                    warnings.Add(string.Format("号機{0}: 本数があるのに金額が0", r.MachineId));
            }
            if (total == 0)
                warnings.Add("全台合計が0円。取得失敗の疑い");

            return warnings;
        }

        private static void WriteTsv(IList<SalesRow> rows, string path)
        {
            using (var w = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                w.WriteLine("検量\t年月\t日付\t号機コード\t号機名\t本数\t総額\t現金\tカード");
                foreach (var r in rows) w.WriteLine(r.ToTsv());
            }
        }

        private static string Summarize(DateTime day, IList<SalesRow> rows, IList<string> warnings)
        {
            int units = 0, total = 0;
            foreach (var r in rows) { units += r.Units; total += r.Total; }

            var sb = new StringBuilder();
            sb.AppendFormat("{0:yyyy-MM-dd} 売上: {1}台 {2}本 {3:N0}円",
                            day, rows.Count, units, total);
            foreach (var w in warnings) sb.AppendLine().Append("⚠ ").Append(w);
            return sb.ToString();
        }

        /// <summary>通知は best-effort。失敗しても本処理の結果は変えない。</summary>
        private static void Notify(Config config, string message)
        {
            if (string.IsNullOrEmpty(config.NotifyWebhookUrl)) return;
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Encoding = Encoding.UTF8;
                    var payload = "{\"content\":" + Quote(message) + "}";
                    client.UploadString(config.NotifyWebhookUrl, payload);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("通知に失敗（本処理は継続）: " + e.Message);
            }
        }

        private static string Quote(string s)
        {
            var sb = new StringBuilder("\"");
            foreach (var ch in s)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.Append('"').ToString();
        }
    }
}
