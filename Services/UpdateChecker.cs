using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace NetworkDiagramApp
{
    public class UpdateChecker
    {
        private const string CURRENT_VERSION = "1.1.0";
        private const string GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest";
        
        // TODO: GitHubのユーザー名とリポジトリ名に置き換えてください
        private const string GITHUB_OWNER = "kynmg22";
        private const string GITHUB_REPO = "network-diagram-tool";

        public class UpdateInfo
        {
            public bool HasUpdate { get; set; }
            public string CurrentVersion { get; set; } = CURRENT_VERSION;
            public string LatestVersion { get; set; } = "";
            public string DownloadUrl { get; set; } = "";
            public string ReleaseNotes { get; set; } = "";
            public string ErrorMessage { get; set; } = "";
        }

        public static async Task<UpdateInfo> CheckForUpdatesAsync()
        {
            var result = new UpdateInfo();

            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", "NetworkDiagramApp");
                client.Timeout = TimeSpan.FromSeconds(5);

                string apiUrl = GITHUB_API_URL
                    .Replace("{owner}", GITHUB_OWNER)
                    .Replace("{repo}", GITHUB_REPO);

                var response = await client.GetAsync(apiUrl);

                if (!response.IsSuccessStatusCode)
                {
                    result.ErrorMessage = $"更新チェック失敗: HTTP {response.StatusCode}";
                    return result;
                }

                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // バージョン番号を取得（tag_nameから "v" を除去）
                string tagName = root.GetProperty("tag_name").GetString() ?? "";
                result.LatestVersion = tagName.TrimStart('v');

                // ダウンロードURL
                result.DownloadUrl = root.GetProperty("html_url").GetString() ?? "";

                // リリースノート
                result.ReleaseNotes = root.GetProperty("body").GetString() ?? "";

                // バージョン比較
                result.HasUpdate = CompareVersions(CURRENT_VERSION, result.LatestVersion) < 0;

                return result;
            }
            catch (HttpRequestException)
            {
                result.ErrorMessage = "インターネット接続を確認してください";
                return result;
            }
            catch (TaskCanceledException)
            {
                result.ErrorMessage = "接続がタイムアウトしました";
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"更新チェックエラー: {ex.Message}";
                return result;
            }
        }

        private static int CompareVersions(string current, string latest)
        {
            try
            {
                var currentParts = current.Split('.');
                var latestParts = latest.Split('.');

                for (int i = 0; i < Math.Max(currentParts.Length, latestParts.Length); i++)
                {
                    int currentNum = i < currentParts.Length ? int.Parse(currentParts[i]) : 0;
                    int latestNum = i < latestParts.Length ? int.Parse(latestParts[i]) : 0;

                    if (currentNum < latestNum) return -1;
                    if (currentNum > latestNum) return 1;
                }

                return 0;
            }
            catch
            {
                return 0; // 比較失敗時は更新なしとみなす
            }
        }

        public static string GetCurrentVersion()
        {
            return CURRENT_VERSION;
        }
    }
}
