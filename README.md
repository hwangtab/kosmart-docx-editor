# 🪄 마법의 AI 워드 표 편집기 (DOCX MCP Server)

안녕하세요! 이 도구는 한국스마트협동조합 실무를 위해 특수 제작된 **DOCX(워드/한글 변환 문서) 전용 AI 플러그인(MCP)** 입니다. 

평소에 AI(클로드, 커서 등)에게 "이 워드 제안서 안의 표 금액을 10만원으로 고쳐줘"라고 시키면, AI가 엉뚱한 코드를 짜서 무한루프(로딩)에 빠지거나 문서 파일을 통째로 깨뜨리는 경험을 해보셨나요? 

이제 이 툴박스를 **복사+붙여넣기 한 번**으로 내 AI에 꽂아주기만 하면, 아무리 복잡하게 **병합된 셀(Merged Cells)** 이 섞여 있는 악랄한 표라도 AI가 단 1초 만에 깔끔하게 읽고, 워드 줄 간격과 폰트 서식을 1mm도 망가뜨리지 않고 빈칸만 쏙쏙 덮어써 줍니다! 😎

---

## 🚀 딱 1분 만에 내 AI에 장착하기 (설치 방법)

파이썬을 전혀 몰라도 괜찮습니다! 터미널(까만 화면)을 열어 아래 **설치 마법사 명령어**를 그냥 복붙하고 엔터만 치세요.

### 🤖 1. 나는 (터미널에서 쓰는) 'Claude Code'를 쓰고 있다면?
아래 코드 네 줄을 순서대로 터미널에 한 줄씩 복사해서 붙여넣고 엔터를 치세요.

```bash
# 1. 파일 다운로드 받기
git clone https://github.com/hwangtab/kosmart-docx-editor.git

# 2. 다운받은 폴더로 쏙 들어가기
cd kosmart-docx-editor

# 3. 도구(파이썬 가상환경) 자동 설치하기
python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt

# 4. 내 클로드 엔진에 이 마법도구 영구적으로 박아넣기 (끝!)
claude mcp add -s global kosmart_docx_editor $(pwd)/venv/bin/python $(pwd)/server.py
```
> 🎉 **성공!** 이제 아무 폴더에서나 `claude`를 켜고 "이 워드 문서 표 제일 마지막 칸을 이걸로 바꿔주라"하면 알아서 다 해줍니다!

---

### 🧑‍🚀 2. 나는 'Antigravity (구글 딥마인드 AI)' 를 쓰고 있다면?

Antigravity는 시스템 내부 설정 파일(`mcp_config.json`)을 열어 코드를 한 블럭 넣어주면 됩니다.
1. 위 Claude Code 설치 안내의 **1~3번 명령어까지 똑같이 터미널에 복사**해서 설치를 마칩니다.
2. 터미널에 `nano ~/.gemini/antigravity/mcp_config.json` 을 입력하여 설정 파일을 엽니다. (없으면 새로 빈 내용을 만듭니다)
3. 아래의 JSON 코드를 문서 안에 붙여넣습니다. `내사용자이름` 은 본인의 맥북 계정명으로 변경해주세요.
   ```json
   {
     "mcpServers": {
       "kosmart_docx_editor": {
         "command": "/Users/내사용자이름/.../kosmart-docx-editor/venv/bin/python",
         "args": [
           "/Users/내사용자이름/.../kosmart-docx-editor/server.py"
         ]
       }
     }
   }
   ```
4. `Ctrl + O`, `Enter`를 눌러 저장하고 `Ctrl + X`로 빠져나옵니다. 설치 끝!

---

### 🟢 2. 나는 'Cursor(커서)'나 'VS Code IDE'를 쓰고 있다면?

먼저, 방금 전 1번 안내에 있는 **1, 2, 3번 명령어까지만 똑같이 터미널에 복사해서 실행(설치)**해주세요.

그 다음, 에디터 설정 창을 열어서 도구를 연결해 주어야 합니다.
1. Cursor 에디터를 열고 **설정(Settings)** 창 열기
2. 왼쪽 메뉴에서 **Features -> MCP** 메뉴 클릭
3. **[+ Add New MCP Server]** 버튼 꾹 누르기
4. 칸을 요렇게 채워주세요.
   * **Name**: 마음대로! (예: `마법의워드수정기`)
   * **Type**: `command` 선택
   * **Command**: (아래 명령어를 주의해서 복붙해주세요. 단, 경로상의 `내사용자이름`을 꼭 본인의 맥 이름으로 바꿔주세요!)
     ```bash
     /Users/내사용자이름/.../kosmart-docx-editor/venv/bin/python /Users/내사용자이름/.../kosmart-docx-editor/server.py
     ```
5. 저장 후 초록색 불(🟢)이 들어오면 성공!

---

## 🛠 어떤 무기들이 들어있나요?

AI가 위 설치 과정을 마치면, 스스로 알아서 아래의 강력한 7가지 마법 스킬을 쓸 수 있게 됩니다.

### 📝 텍스트 조작 (안전 덮어쓰기)
* **`safe_replace_docx_text_keyword` (무결점 단어 치환)**: 문서 본문, 표, **머리글(Header), 바닥글(Footer)**은 물론 파란색 **하이퍼링크** 텍스트까지 찾아내어 원래 지정된 색상, 굵기 등의 서식을 100% 보존하면서 글씨만 싹 바꿔줍니다.
* **`safe_replace_docx_cell` (안전 셀 덮어쓰기)**: 워드 안에 숨어있는 보이지 않는 태그 서식을 보존하면서 내용물 텍스트만 덮어써 줍니다.
* **`safe_replace_text_box_keyword` (텍스트 상자 해킹)**: `python-docx`가 절대 못 읽는 **떠 있는 도형(Floating Shapes)이나 텍스트 상자** 내부의 글씨를, 원본 파일의 XML 압축을 실시간으로 해제하고 투입하여 강제로 바꿔버리는 사기급 기능입니다.

### 🖼️ 이미지 제어
* **`list_docx_images`**: 문서 내에 짱박혀 있는 증명사진 등 인라인 이미지들의 (설정된 크기와) 고유 번호를 싹 뽑아줍니다.
* **`replace_docx_image`**: 문서의 레이아웃, 크기 여백을 1mm도 어긋나지 않게 유지하면서 기존 템플릿 이미지를 내가 원하는 새로운 이미지 파일로 감쪽같이 교체합니다.

### 📊 표(Table) & 리스트 제어
* **`read_docx_table_flat_index` (워드 투시경)**: 아무리 복잡한 표라도 눈에 보이는 네모 칸(Cell) 개수대로 순서를 매겨서 깔끔하게 글자를 읽어줍니다. 엑셀 지옥의 '병합된 셀' 에러 걱정 끝!
* **`read_docx_table_optimized` (대용량 초고속 스캐너)**: 수백 페이지 엑셀 변환 파일도 메모리 에러(OOM) 없이, 0.1초 만에 백그라운드에서 XML만 스트리밍으로 긁어옵니다. 줄바꿈(`\n`)도 보존합니다.
* **`append_docx_list_item` (리스트 증식)**: 기존 `1. 목표` 같은 번호 매기기 리스트 아래에 `2. 전략`을 추가할 때, 엉뚱한 폰트가 적용되지 않도록 원본 부모의 XML 서식을 딥카피하여 완벽하게 리스트를 늘려줍니다.

---
**License**: 자유롭게 복사하고 회사에서 마구마구 동네방네 퍼트려 써주세요! (MIT License)
