import { defineConfig } from "vite";

// GitHub Pages 배포 시, 정적 자산 경로가 /REPO_NAME/ 아래로 잡혀야 흰 화면(404)이 안 납니다.
export default defineConfig({
  base: "file-lister",
});