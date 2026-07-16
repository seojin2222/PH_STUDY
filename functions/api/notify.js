// Cloudflare Pages Function — POST /api/notify
// 출결 이벤트가 발생했을 때 브라우저가 이 엔드포인트를 호출하면,
// 여기(서버)에서만 솔라피 API 키를 사용해 카카오 알림톡을 발송한다.
// API 키/시크릿/템플릿ID는 전부 Cloudflare 환경변수에서 읽어오며 클라이언트 번들에는 절대 포함되지 않는다.

const TEMPLATE_ENV_KEY = {
  등원: "KAKAO_TPL_ARRIVAL",
  외출: "KAKAO_TPL_OUT",
  외출복귀: "KAKAO_TPL_RETURN",
  하원: "KAKAO_TPL_LEAVE",
};

export async function onRequestPost({ request, env }) {
  try {
    const { phone, type, name, time, reason, expectedReturn } = await request.json();

    if (!phone || !type || !name || !time) {
      return jsonResponse({ ok: false, error: "필수 값 누락 (phone, type, name, time)" }, 400);
    }
    const templateEnvKey = TEMPLATE_ENV_KEY[type];
    if (!templateEnvKey) {
      return jsonResponse({ ok: false, error: `알 수 없는 유형: ${type}` }, 400);
    }
    const templateId = env[templateEnvKey];
    if (!env.SOLAPI_API_KEY || !env.SOLAPI_API_SECRET || !env.KAKAO_PFID || !env.KAKAO_SENDER || !templateId) {
      return jsonResponse({ ok: false, error: "서버에 알림톡 설정(환경변수)이 완료되지 않았어요" }, 500);
    }

    const variables = { "#{이름}": name, "#{시간}": time };
    if (type === "외출") {
      variables["#{사유}"] = reason || "기타";
      variables["#{복귀시간}"] = expectedReturn || "미정";
    }

    const authHeader = await buildSolapiAuthHeader(env.SOLAPI_API_KEY, env.SOLAPI_API_SECRET);

    const solapiRes = await fetch("https://api.solapi.com/messages/v4/send", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: authHeader },
      body: JSON.stringify({
        message: {
          to: onlyDigits(phone),
          from: onlyDigits(env.KAKAO_SENDER),
          kakaoOptions: {
            pfId: env.KAKAO_PFID,
            templateId,
            variables,
            disableSms: true, // 알림톡 실패 시 문자 대체발송 안 함 (비용 통제)
          },
        },
      }),
    });

    const data = await solapiRes.json();
    return jsonResponse({ ok: solapiRes.ok, data }, solapiRes.ok ? 200 : 502);
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) }, 500);
  }
}

function onlyDigits(s) {
  return String(s || "").replace(/\D/g, "");
}

function jsonResponse(obj, status) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
}

/* 솔라피 HMAC-SHA256 인증 헤더 생성
   Authorization: HMAC-SHA256 apiKey=..., date=..., salt=..., signature=HmacSHA256(date+salt, apiSecret) */
async function buildSolapiAuthHeader(apiKey, apiSecret) {
  const date = new Date().toISOString();
  const salt = crypto.randomUUID().replace(/-/g, "");
  const signature = await hmacSha256Hex(apiSecret, date + salt);
  return `HMAC-SHA256 apiKey=${apiKey}, date=${date}, salt=${salt}, signature=${signature}`;
}

async function hmacSha256Hex(secret, message) {
  const enc = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  const sigBuf = await crypto.subtle.sign("HMAC", key, enc.encode(message));
  return [...new Uint8Array(sigBuf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
