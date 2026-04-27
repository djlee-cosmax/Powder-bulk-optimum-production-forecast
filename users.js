// 등록된 사용자 목록 (초기 비밀번호 해시)
// 비밀번호 변경 시 localStorage에 새 해시가 저장됨 (PC별로 별도)
// 초기 비밀번호로 되돌리려면 localStorage의 pwd_override_<id>를 삭제하면 됨
//
// 사용자 추가 절차:
//   1) WSL에서 node로 PBKDF2 해시 생성
//      node -e "const c=require('crypto');const s=c.randomBytes(16).toString('hex');const h=c.pbkdf2Sync('초기비밀번호', Buffer.from(s,'hex'), 100000, 32, 'sha256').toString('hex');console.log({salt:s,hash:h})"
//   2) 아래 USERS 객체에 항목 추가 (이름 + 소속 + salt/hash)
//      "ID": { name: "홍길동", dept: "생산3팀", salt: "...", hash: "...", iterations: 100000 }
//   3) git add users.js && git commit && git push
//
// 소속(dept)이 없으면 이름만 표시되며, 있으면 "생산3팀 이동준" 형태로 표시됩니다.

window.USERS = {
  "djlee": {
    name: "이동준",
    dept: "생산3팀",
    admin: true,
    salt: "98b8c3c75b30f6e3b58278a2e7bcc13c",
    hash: "6110f78c92d158d2cf13f0de8feb1318c56333a5e24168c15c609c43e461ed11",
    iterations: 100000
  },
  "dhwon": {
    name: "원대한",
    dept: "생산운영팀",
    salt: "9212a000038bdcfe18964ac6aa054d4b",
    hash: "a733a2e7d50a39f0e1d9a126c0e0abd2d6033211c2021d4b161c23b2a5314941",
    iterations: 100000
  }
};
