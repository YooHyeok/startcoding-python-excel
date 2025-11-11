# 파이썬 엑셀 라이브러리 종류
- xlwings  
- pandas  
- openPyXl
- Pywin32

액셀 자체기능을 최대한 활용할 수 있는 xlwings와 데이터 처리 및 분석에 강점이 있는 pandas 두 라이브러리를 조합하여 진행

# xlwings
<details>
<summary>접기/펼치기</summary>
<br>


파이썬 액셀 자동화 라이브러리의 한 종류

1. 액셀을 직접 제어할 수 있기 때문에 액셀 자체 기능을 최대한 활용 가능.
2. 액셀을 활용해서 작업을 하기 때문에 액셀 설치 필요
3. DRM 우회가 가능

### 자동화 가능 항목
- 파일 시트 생성, 수정 저장
- 셀 데이터 추가, 수정 삭제
- 행 생성, 삭제
- 스타일 변경
- 취합하기
- 복사, 붙여넣기
- 병합, 병합해제
- 수식 입력, PDF 변환
- 대용량 데이터 처리 및 분석, 그래프 시각화

## 액셀의 구성요소
- 프로그램
- 워크북
- 워크시트
- 셀

`프로그램(1) > 워크북(N) > 워크시트(N) > 셀(N)`

하나의 프로그램 안에는 여러개의 워크 북이 들어있을 수 있다.  
(여기서 워크 북은 하나의 액셀 파일을 말한다.)  
하나의 워크북 안에는 여러개의 워크시트가 들어있을 수 있다.  
하나의 워크시트 안에는 여러개의 셀이 들어있다.  

## xlwings의 구성요소
- app - 프로그램
- book - 워크북
- sheet - 워크시트
- range - 셀 범위

`app(1) > book(N) > sheet(N) > range(N)`

하나의 app 안에는 여러개의 book이 들어있을 수 있다.  
하나의 book 안에는 여러개의 sheet가 들어있을 수 있다.  
하나의 sheet에서 range를 가져올 수 있다.  
(여기서 range는 하나의 셀 또는 여러개의 셀이 될 수 있다.)  

### 워크북 다루기 명령어
| 명령어                          | 기능                |
| ------------------------------ | :----------------- |
| app = xw.App(add_book=False)	 | 액셀 앱 만들기       |
| wb = app.books.add()           | 액셀 워크북 생성하기  |
| wb = app.books.open("파일경로)  | 기존 워크북 불러오기  |
| wb.save("파일경로)              | 다른 이름으로 저장    |
| wb.save()                      | 저장하기            |
| app.quit()                     | 액셀 앱 닫기         |


### 워크시트 다루기 명령어
| 명령어                          | 기능                |
| ------------------------------ | :------------------ |
| wb.sheets.books.add("이름")    | 새로운 시트 생성하기    |
| ws = wb.sheets["이름"]         | 이름으로 시트 생성하기   |
| ws = wb.sheets[0]             | 인덱스로 시트 생성하기   |
| ws.name = "변경할 이름"         | 시트 이름 변경         |
| wb.sheets["이름"].delete()     | 시트 삭제             |
| wb.sheets["이름"].activate()   | 시트 활성화           |
| wb.sheets["이름"].clear()      | 시트 내용 전체 삭제    |

위 명령어들은 외워서 사용하지 않고 필요할때 찾아 쓰도록 한다.  


### 셀 다루기 명령어
| 명령어                             | 기능                |
| --------------------------------- | :------------------ |
| wb.range('A1').value = '데이터'    | A1 셀에 데이터 입력   |
| wb.range('A1')                    | 셀 한개 접근하기      |
| wb.range('A1').value              | 셀 한개 값 가져오기   |
| wb.range('A1:D9')                 | 셀 여러개 접근하기    |
| wb.range('A1').expand('table')    | 셀 확장해서 선택하기  |

# xlwings 라이브러리 install
```bash
pip install xlwings
```

</details>
<br>


# 액셀 자동화 기초 
## 액셀 파일 다루기
<details>
<summary>접기/ 펼치기</summary>

### 액셀 신규 생성(앱,워크북 생성, 시트명 변경, 액셀 저장 및 종료)
```py
import xlwings as xw # xlwinngs import 및 xw 별칭 부여

# 액셀 앱 만들기
app = xw.App(add_book=False) # 액셀 생성 및 오픈

# 액셀 워크북 만들기
wb = app.books.add() # 생성되어 열린 액셀파일에 워크시트(통합문서 2) 생성

# 워크시트 선택(이름)
ws = wb.sheets['Sheet1']

# 시트 이름 변경
ws.name = "영업1팀"

# 다른 이름으로 액셀 저장
wb.save('교육이슈현황.xlsx')

# 워크북 닫기
wb.close()

# 액셀 앱 닫기
app.quit()
```

### 액셀 불러오기(시트 추가 및 순서 제어, 시트 삭제 및 활성화)
```py
import xlwings as xw # xlwinngs import 및 xw 별칭 부여

# 액셀 앱 만들기
app = xw.App(add_book=False) # 액셀 생성 및 오픈

# 액셀 워크북 불러오기
wb = app.books.open('교육이슈현황.xlsx')

# 새로운 시트 생성
wb.sheets.add('영업2팀') # [`영업2팀` 영업1팀]
wb.sheets.add('마케팅팀') # [`마케팅팀` 영업2팀 영업1팀]  

# 시트 생성 - [마케팅팀 영업2팀 영업1팀]에서 마케팅팀~영업2팀 사이에 생성
wb.sheets.add(name='세무회계팀', before='영업2팀') # [마케팅팀 `세무회계팀` 영업2팀 영업1팀]

# 시트 생성 - [마케팅팀 세무회계팀 영업2팀 영업1팀]에서 영업2팀~영업1팀 사이에 생성
wb.sheets.add(name='영업3팀', after='영업2팀') # [마케팅팀 세무회계팀 영업2팀 `영업3팀` 영업1팀]

# 워크 시트 선택 (인덱스)
wb.sheets[0] # <Sheet [교육이슈현황.xlsx]마케팅팀>
wb.sheets[3] # <Sheet [교육이슈현황.xlsx]영업3팀>

# 워크 시트 선택 (이름)
wb.sheets['영업3팀']

# 시트 삭제
wb.sheets['영업3팀'].delete() # [마케팅팀 세무회계팀 영업2팀 영업1팀]

# 시트 활성화 - activate()
wb.sheets['마케팅팀'].activate()

# 액셀 저장 및 종료
wb.save()
app.quit()
```

</details>
<br>
<hr>

## 셀 다루기
<details>
<summary>접기/ 펼치기</summary>

### 액셀 데이터 삽입
- 마케팅 팀의 교육 이수 현황 로직 구현
  ```py
  import xlwings as xw

  app = xw.App(add_book=False) # 액셀 프로그램(앱) 생성
  wb = app.books.open('교육이수현황.xlsx') # 교육이슈현황.xlsx 액셀파일 열기 및 변수 저장
  ws = wb.sheets['마케팅팀'] # 열린 액셀파일의 마케팅팀 시트 선택 및 변수 저장

  # 데이터 1개 입력 - A1셀에 '성명', B1셀에 '1월' 값 삽입
  ws.range('A1').value = '성명'
  ws.range('B1').value = '1월'

  # A2, B2 각각의 셀에 리신, 12 값 삽입
  ws.range('A2').value = '리신'
  ws.range('B2').value = 12

  # 데이터 여러 개 입력 - 리스트
  ws.range('A3').value = ['이즈리얼', 5]
  ws.range('A4').value = [['야스오', 0], ['요네', 0]] # A4 셀 기준 2차원 배열 형태로 데이터 삽입 (A4, B4, A5, B5열 순차적으로 삽입됨)

  # 데이터 여러 개 입력 - 세로
  ws.range('A6').options(transpose=True).value = ['피즈', '샤코'] # A6기준으로 세로 리스트 삽입
  ws.range('B6').options(transpose=True).value = [15.5, 8] # B6기준으로 세로 리스트 삽입

  # 수식 입력
  ws.range('A8').value = '합계'
  ws.range('B8').value = '=sum(B2:B7)'
  ```

크롤릴을 통해 실시간으로 수집한 데이터를 원하는 위치에 넣을 수도 있으며,  
다른 파일, 데이터 베이스에서 추출한 데이터를 넣을수도 있다.  


</details>
<br>
<hr>
