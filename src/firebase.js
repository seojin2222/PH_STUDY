// firebase.js
// 1) https://console.firebase.google.com 에서 프로젝트 생성
// 2) 프로젝트 설정 > 앱 추가(웹) 에서 아래 config 복사해서 붙여넣기
// 3) Firestore Database 사용 설정 (프로덕션/테스트 모드)

import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyDL38IdKfV6Zj16nmVdk3lkV_f4CjKWrIU",
  authDomain: "pharos-study.firebaseapp.com",
  projectId: "pharos-study",
  storageBucket: "pharos-study.firebasestorage.app",
  messagingSenderId: "751696720777",
  appId: "1:751696720777:web:44cdde773714d10b1fa758",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
