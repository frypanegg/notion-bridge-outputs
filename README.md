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
| [철강 제조업의 HR 업무를 총괄하는 CHO](sessions/%EC%B2%A0%EA%B0%95-%EC%A0%9C%EC%A1%B0%EC%97%85%EC%9D%98-HR-%EC%97%85%EB%AC%B4%EB%A5%BC-%EC%B4%9D%EA%B4%84%ED%95%98%EB%8A%94-CHO-3db455aa/README.md) | 5 | 해당 자료를 CHO에게 메일로 보고드리고자 해. 내일 13:30 화상으로 대면 보고드리겠다고 하고. 간략하게 본문에 어떤 방향으로 검토했는지 … | 2026. 9. 15. 오후 3:46 |
| [test](sessions/test-3da455aa/README.md) | 1 | 아직 이런 코딩 방식이 어려워. 설명서를 html로 쉽게 보기좋게 만들어서 구글드라이브에 업로드하고링크 보여줘. | 2026. 9. 14. 오전 9:17 |
