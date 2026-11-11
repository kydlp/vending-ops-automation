using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

// フラットなJSONだけを読む最小の設定リーダ。
// 追加アセンブリへの依存を避けたいだけなので、汎用のJSONパーサは使っていない
// （標準ライブラリだけで完結させ、配布物を実行ファイル＋configに保つため）。
//
// 探索順：実行ファイルの位置から親方向へ4階層まで config.local.json を探す。
// 実値を含む config.local.json はリポジトリに含めない（config.example.json のみ）。
internal sealed class Config
{
    public string Endpoint = "";
    public string TenantId = "";
    public string LoginId = "";
    public string Password = "";
    public string NotifyWebhookUrl = "";

    public static Config Load()
    {
        var path = Resolve("config.local.json");
        if (path == null)
            throw new FileNotFoundException(
                "config.local.json が見つかりません（プロジェクトルートに配置してください）。");

        var json = File.ReadAllText(path, Encoding.UTF8);
        var c = new Config();
        c.Endpoint = Str(json, "endpoint") ?? c.Endpoint;
        c.TenantId = Str(json, "tenantId") ?? c.TenantId;
        c.LoginId = Str(json, "loginId") ?? c.LoginId;
        c.Password = Str(json, "password") ?? c.Password;
        c.NotifyWebhookUrl = Str(json, "notifyWebhookUrl") ?? c.NotifyWebhookUrl;
        return c;
    }

    private static string Resolve(string file)
    {
        var dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        for (int i = 0; i < 4 && dir != null; i++)
        {
            var p = Path.Combine(dir, file);
            if (File.Exists(p)) return p;
            dir = Path.GetDirectoryName(dir);
        }
        return null;
    }

    private static string Str(string json, string key)
    {
        var m = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
        if (m.Success) return Regex.Unescape(m.Groups[1].Value);
        var n = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*(-?\\d+)");
        return n.Success ? n.Groups[1].Value : null;
    }
}
