// Cloudflare Worker — 代理触发 GitHub Actions
// 环境变量: GITHUB_TOKEN (在 Cloudflare Dashboard 的 Workers & Pages > 你的 Worker > Settings > Variables 里添加)

const REPO   = "yingchuang2894-star/sheet-dashboard";
const WORKFLOW = "update_data.yml";
const BRANCH = "main";
const COOLDOWN_MS = 5 * 60 * 1000; // 5 分钟内只能触发一次

export default {
  async fetch(request, env) {
    const cors = {
      "Access-Control-Allow-Origin": "https://yingchuang2894-star.github.io",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Max-Age": "86400",
    };

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: cors });
    }
    if (request.method !== "POST") {
      return new Response("Method not allowed", { status: 405, headers: cors });
    }

    // 简单冷却：用 KV 或直接依赖客户端控制（此处信任前端 cooldown）
    const token = env.GITHUB_TOKEN;
    if (!token) {
      return new Response(JSON.stringify({ ok: false, msg: "未配置 GITHUB_TOKEN" }), {
        status: 500, headers: { ...cors, "Content-Type": "application/json" }
      });
    }

    const apiResp = await fetch(
      `https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/dispatches`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "Content-Type": "application/json",
          "User-Agent": "sheet-dashboard-worker",
        },
        body: JSON.stringify({ ref: BRANCH }),
      }
    );

    const ok = apiResp.status === 204; // 成功时 GitHub 返回 204
    return new Response(
      JSON.stringify({ ok, status: apiResp.status }),
      { headers: { ...cors, "Content-Type": "application/json" } }
    );
  },
};
