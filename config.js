//  Supabase 접속 정보.
//  Supabase → Project Settings → API Keys 에서 복사합니다.
//   · URL  : Project URL
//   · key  : Publishable key (sb_publishable_…) 또는 예전 형식의 anon public 키
//
//  ※ 이 키는 **공개되어도 되는 값**입니다. 브라우저에 그대로 실려야 동작합니다.
//     실제 접근 통제는 setup.sql 의 RLS 정책(로그인한 사람만 읽기)이 합니다.
//     Secret key(sb_secret_…) · service_role 키는 **절대** 여기 넣지 마십시오 — RLS 를 건너뜁니다.
//
//  ※ 값에 꺾쇠 < > 를 남기지 마십시오. 따옴표 안에 값만 있어야 합니다.
window.SUPA = {
  url: "https://ryrhtvaeutqejdekpari.supabase.co",
  key: "sb_publishable_DBRKg4svB4CBjiMptQTD2Q_ywYTTYIL",
};
