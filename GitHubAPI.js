// ============================================
// GitHubAPI.js — GitHub REST API ラッパー
// バグ報告から自動PR作成 (BugReportHandler.js から呼ばれる)
// ============================================

var GITHUB_REPO_OWNER = 'eisei-strong';
var GITHUB_REPO_NAME = 'eisei';

function getGithubPat_() {
  return PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
}

/** GitHub API 共通リクエスト */
function githubRequest_(method, path, payload) {
  var pat = getGithubPat_();
  if (!pat) throw new Error('GITHUB_PAT が Script Properties に未設定');
  var url = 'https://api.github.com' + path;
  var options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + pat,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code >= 400) {
    throw new Error('GitHub API ' + method + ' ' + path + ' failed (' + code + '): ' + body);
  }
  return body ? JSON.parse(body) : null;
}

/**
 * リポジトリのファイル取得
 * @return {{content: string, sha: string, path: string}}
 */
function githubGetFile_(filePath, ref) {
  var path = '/repos/' + GITHUB_REPO_OWNER + '/' + GITHUB_REPO_NAME + '/contents/' + encodeURIComponent(filePath);
  if (ref) path += '?ref=' + encodeURIComponent(ref);
  var data = githubRequest_('get', path);
  return {
    content: Utilities.newBlob(Utilities.base64Decode(data.content)).getDataAsString(),
    sha: data.sha,
    path: data.path
  };
}

/** 指定ブランチの最新コミットSHA取得 */
function githubGetBranchSha_(branch) {
  var data = githubRequest_('get', '/repos/' + GITHUB_REPO_OWNER + '/' + GITHUB_REPO_NAME + '/git/refs/heads/' + branch);
  return data.object.sha;
}

/** ブランチ作成（既存ブランチからコピー） */
function githubCreateBranch_(newBranch, fromBranch) {
  var sha = githubGetBranchSha_(fromBranch || 'main');
  return githubRequest_('post', '/repos/' + GITHUB_REPO_OWNER + '/' + GITHUB_REPO_NAME + '/git/refs', {
    ref: 'refs/heads/' + newBranch,
    sha: sha
  });
}

/** ファイル更新（既存sha必須） */
function githubUpdateFile_(filePath, newContent, commitMessage, branch) {
  // 既存ファイルのsha取得
  var existing = githubGetFile_(filePath, branch);
  var encoded = Utilities.base64Encode(Utilities.newBlob(newContent).getBytes());
  return githubRequest_('put', '/repos/' + GITHUB_REPO_OWNER + '/' + GITHUB_REPO_NAME + '/contents/' + encodeURIComponent(filePath), {
    message: commitMessage,
    content: encoded,
    sha: existing.sha,
    branch: branch
  });
}

/** PR作成 */
function githubCreatePR_(title, body, headBranch, baseBranch) {
  return githubRequest_('post', '/repos/' + GITHUB_REPO_OWNER + '/' + GITHUB_REPO_NAME + '/pulls', {
    title: title,
    body: body,
    head: headBranch,
    base: baseBranch || 'main'
  });
}

/** GASエディタからの動作確認用 */
function testGithubApi() {
  try {
    var sha = githubGetBranchSha_('main');
    Logger.log('main HEAD sha: ' + sha);
    var file = githubGetFile_('PostApp.js');
    Logger.log('PostApp.js sha: ' + file.sha);
    Logger.log('PostApp.js 1行目: ' + file.content.split('\n')[0]);
    Logger.log('✅ GitHub API 接続OK');
  } catch (e) {
    Logger.log('❌ GitHub API エラー: ' + e.message);
  }
}
