# notion-bridge-outputs

노션 세션에서 맥미니 Claude Code에 요청한 결과물이 턴 단위로 쌓이는 저장소입니다. 이 파일과 각 README는 브릿지가 자동으로 생성합니다.

```
sessions/<세션>-<노션페이지ID>/
  README.md                 세션의 턴 목록
  turns/<턴번호>-<날짜>/
    README.md               요청 · 요약 · 결과물 목록
    files/                  결과물 (원래 이름과 폴더 구조 유지)
    changes.patch           프로젝트 코드 변경분
```

100MB가 넘는 파일은 git 대신 [Releases](https://github.com/frypanegg/notion-bridge-outputs/releases)에 첨부되고 턴 README에서 링크됩니다.

## 세션

| 세션 | 턴 | 마지막 요청 | 갱신 |
| --- | --- | --- | --- |
| [test](sessions/test-3da455aa/README.md) | 1 | 아직 이런 코딩 방식이 어려워. 설명서를 html로 쉽게 보기좋게 만들어서 구글드라이브에 업로드하고링크 보여줘. | 2026. 9. 14. 오전 9:17 |
