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

### 액셀 데이터 가져오기

```py
import xlwings as xw

app = xw.App(add_book=False)
wb = app.books.open('교육이슈현황.xlsx')
ws = wb.sheets['마케팅팀']

# 셀 한개 접근하기
ws.range('A1') # <Range [교육이슈현황.xlsx]마케팅팀!$A$1>

# 셀 한계 값 가져오기
ws.range('A1').value # '성명'

# 셀 여러개 접근하기
ws.range('A1:B8') # <Range [교육이슈현황.xlsx]마케팅팀!$A$1:$B$8>

# 셀 여러개 값 가져오기
ws.range('A1:B8').value # 2차원 리스트 값 가져오게 됨

# C1 ~ D1에 신규 월 추가
ws.range('C1').value = ['2월', '3월']
```

### 복사 붙여넣기 - xlwings에서만 사용할 수 있는 특별한 기능
```py
import xlwings as xw

app = xw.App(add_book=False)
wb = app.books.open('교육이슈현황.xlsx')
ws = wb.sheets['마케팅팀']

# C1 ~ D1에 신규 월 추가
ws.range('C1').value = ['2월', '3월']

# 복사
ws.range('B2:B8').copy()

# 붙여넣기
ws.range('C2').paste()
ws.range('D2').paste()
```

### 셀 확장해서 선택하기
데이터가 없는 곳이 나오면 끊긴다.
```py

# 아래방향
ws.range('A1').expand('down').value
# 우측방향
ws.range('A1').expand('right').value
# 테이블(2차원 배열)
ws.range('A1').expand('table').value
```

액셀에서 원하는 부분의 데이터를 가지고 올 수 있는 것은 굉장히 활용도가 높은 작업이다.  
데이터를 다른 파일에 넣을 수 있으며, 웹사이트 자동화나 데이터 분석, 시작화 등에 활용 가능하다.

### 폰트
```py
import xlwings as xw

app = xw.App(add_book=False)
wb = app.books.open('교육이슈현황.xlsx')
ws = wb.sheets['마케팅팀']

# 폰트 사이즈 변경
ws.range('A1:D1').font.size = 13

# 폰트 스타일 bold
ws.range('A1:D1').font.bold = True

# 셀 배경 및 폰트 색 변경
ws.range('A1:D1').color = (255, 125, 0) # 오렌지 색상
ws.range('A1:D1').font.color = (250, 250, 0) # 밝은 회색
```

</details>
<br>
<hr>

# 폴더 파일 자동화

## 상대경로 vs 절대경로
<details>
<summary>접기/펼치기</summary>
<br>

### 경로(path)
파일 및 폴더의 위치를 말한다.
ex) `Documents/images/프사/셀카.jpg`

### 디렉토리
폴더를 디렉토리라고도 표현한다.
ex) `Documents/images/프사/`

### 절대경로
최초 디렉토리를 기준으로 경로를 설정한 것
ex) `C:/Documents/images/프사/셀카.jpg`

### 상대경로(relative path)
현재 디렉토리를 기준으로 경로를 설정한 것
ex) 현재위치: `C:/Documents/images` → 상대경로:`./프사/셀카.jpg`

### OS 모듈 (Operating System)
파이썬에 기본적으로 내장되어 있는 라이브러리로 운영체제에서 제공하는 여러기능을 파이썬에서 사용할 수 있게 해준다.

#### os 기본 명령어
  | 명령어                          | 기능                   |
  | ------------------------------ | :--------------------- |
  | os.mkdir('path')            	 | 디렉토리 만들기          |
  | os.rmdir('path')            	 | 디렉토리 만들기          |
  | os.rename('path1','path2')     | 파일 이름 변경 또는 이동  |
  | os.remove('path')            	 | 파일 삭제               |
  | os.getcwd()            	       | 현재 경로 확인          |


#### os.path 모듈 명령어
  | 명령어                            | 기능                                |
  | -------------------------------- | :---------------------------------- |
  | os.path.exists('path')        	 | 파일 및 디렉토리의 존재여부 확인<br>(반환값은 bool형) |
  | os.path.join('path1', 'path2')   | 경로 합치기                          |
  | os.path.splitext('path')         | 파일명과 확장자를 분리해서 반환         |
  | os.path.split('path')            | 디렉토리와 파일로 분리해서 반환         |
  | os.path.basename('path')         | 경로의 기본 이름을 반환(ex, test.xlsx) |
</details>
<br>
<hr>

## 폴더 만들기
<details>
<summary>접기/펼치기</summary>
<br>


### 상대 경로 이용1
```py
import os # os의 경우 내장 라이브러리 이므로 설치하지 않아도 됨
os.mkdir('참이슬') # mkdir: make directory - 디렉토리 생성 (./참이슬 에서 ./ 생략)

```
### 상대 경로 이용2
```py
import os
os.mkdir('참이슬/후레시')

```
### 상대 경로 이용3
```py
import os
os.mkdir('../카스') # 현재 디렉토리 기준 상위폴더에 카스 디렉토리 생성

```
### 절대 경로 이용1
```py
import os
os.mkdir('c:/Users/프로젝트폴더/02.폴더및파일관리/테라')
```

### 절대 경로 이용2 - escape raw string
```py
import os
# os.mkdir('C:\Users\프로젝트폴더\02.폴더및파일관리\테라') # 오류 발생

""" 
SyntaxError: (unicode error) 'unicodeescape' codec can't decode bytes in position 2-3: truncated \UXXXXXXXX escape
슬래시와 역슬래시는 경로를 구분할 때 같이 사용할 수 있다.  
그러나 역슬래시는 문자열 안에서 튻구한 문자들을 나타낼 때 사용하는 기호이다.  
\n은 줄바꿈, \t는 탭, \0은 null문자  
즉, \02.~ 부분에서 \0이  null문자로 치환되어버렸기 때문에 경로가 제대로 설정되지 않은것이다.  
이때 경로를 복사할 때는 문자의 시작기호 앞에 r 을 붙혀주면 해결이된다.
"""

# escape raw string
os.rmdir(r'C:\Users\프로젝트폴더\02.폴더및파일관리\테라') # 삭제
os.mkdir(r'C:\Users\프로젝트폴더\02.폴더및파일관리\테라')
```

### 폴더가 없을 때만 만들기
```py
import os

path = r'C:\Users\프로젝트폴더\02.폴더및파일관리\테라'

if not os.path.exists(path):
  os.mkdir(path)
```

</details>
<br>
<hr>

## 파일 100개 이름 변경
<details>
<summary>접기/펼치기</summary>
<br>

### 파일 입출력 - txt 파일 생성
```py
f = open('일기.txt', 'w', encoding = 'utf-8')
f.write('오늘 한잔하고 싶은 날이다...')
f.close()
```

### 파일 100개 생성
```py
import random
for i in range(1, 101):
  f = open(f'참이슬/판매_데이터{i}.txt', 'w', encoding = 'utf-8')
  f.write(f'{random.randint(1, 100)}병')
  f.close()
```

### 파일 100개 이름 수정
```py
import os
for i in range(1, 101):
  os.rename(f'참이슬/판매_데이터{i}.txt', f'참이슬/sale_data{i}.txt')
```
</details>
<br>
<hr>

## 파일 목록 추출하기
<details>
<summary>접기/펼치기</summary>
<br>

### glob 모듈
폴더 안에서 내가 원하는 파일들만 뽑아오고(이동 or 복사) 싶을 때 사용
내가 원하는 액셀 파일들만 가져와서 자동화 프로그램을 동작시킬 수 있다.  
몇가지 기호를 이용하여 패턴을 만들고 경로와 매칭시킨다.
  |  기호           | 기능                                   |
  | -------------- | :------------------------------------- |
  | ?            	 | 문자의 종류와 상관없이 정확히 한글자와 매칭  |
  | *           	 | 길이와 상관없이 모든 문자열과 매칭          |
  | []             | []안의 특정문자 한글자를 의미              |

### shutil 모듈
파일과 폴더를 이동하거나 복사하고 싶을 때 사용
  |  기호                                   | 기능                   |
  | -------------------------------------- | :--------------------- |
  | shutil.move('path1','path2')        	 | 파일 또는 폴더를 이      |
  | shutil.copy('file_path','dir_path') 	 | 파일을 특정폴더로 복사    |

</details>
<br>
<hr>

### glob 사용법
<details>
<summary>접기/펼치기</summary>
<br>

파일 100개 이름 변경에서 구성한 `참이슬/*.txt` 파일들에 코드를 적용한다.

#### glob(): 파일 경로 반환
```py
import glob
glob.glob('참이슬/sale_data1.txt') # ['참이슬/sale_data1.txt'] - glob.glob()는 리스트 형태로 파일의 경로를 반환한다.
```

#### ? 기호
문자의 종류와 상관없이 정확히 한글자와 매칭
```py
import glob
glob.glob('참이슬/sale_data?.txt')
```
##### [출력결과]
```text/plain
['참이슬\\sale_data1.txt',
'참이슬\\sale_data2.txt',
'참이슬\\sale_data3.txt',
'참이슬\\sale_data4.txt',
'참이슬\\sale_data5.txt',
'참이슬\\sale_data6.txt',
'참이슬\\sale_data7.txt',
'참이슬\\sale_data8.txt',
'참이슬\\sale_data9.txt']
```


#### * 기호
길이와 상관없이 모든 문자열과 매칭
```py
import glob
glob.glob('참이슬/*.txt')
```

##### [출력결과]
```text/plain
['참이슬\\sale_data1.txt',
 '참이슬\\sale_data10.txt',
 '참이슬\\sale_data100.txt',
 '참이슬\\sale_data11.txt',
 '참이슬\\sale_data12.txt',
 '참이슬\\sale_data13.txt',
 '참이슬\\sale_data14.txt',
 '참이슬\\sale_data15.txt',
 '참이슬\\sale_data16.txt',
 '참이슬\\sale_data17.txt',
 '참이슬\\sale_data18.txt',
 '참이슬\\sale_data19.txt',
 '참이슬\\sale_data2.txt',
 '참이슬\\sale_data20.txt',
 '참이슬\\sale_data21.txt',
 '참이슬\\sale_data22.txt',
 '참이슬\\sale_data23.txt',
 '참이슬\\sale_data24.txt',
 '참이슬\\sale_data25.txt',
 '참이슬\\sale_data26.txt',
 '참이슬\\sale_data27.txt',
 '참이슬\\sale_data28.txt',
 '참이슬\\sale_data29.txt',
 '참이슬\\sale_data3.txt',
 '참이슬\\sale_data30.txt',
...
 '참이슬\\sale_data95.txt',
 '참이슬\\sale_data96.txt',
 '참이슬\\sale_data97.txt',
 '참이슬\\sale_data98.txt',
 '참이슬\\sale_data99.txt']
```

#### [] 기호
[aeiou], [0-9] (대)괄호안에 있는 문자 하나와 매칭

##### 01. [0-9]
```py
import glob
glob.glob('참이슬/sale_data[0-9].txt')
```
##### [출력결과]
```text/plain
['참이슬\\sale_data1.txt',
 '참이슬\\sale_data2.txt',
 '참이슬\\sale_data3.txt',
 '참이슬\\sale_data4.txt',
 '참이슬\\sale_data5.txt',
 '참이슬\\sale_data6.txt',
 '참이슬\\sale_data7.txt',
 '참이슬\\sale_data8.txt',
 '참이슬\\sale_data9.txt']
```

##### 02. [123]
```py
import glob
glob.glob('참이슬/sale_data[123].txt')
```
##### [출력결과]
```text/plain
['참이슬\\sale_data1.txt', '참이슬\\sale_data2.txt', '참이슬\\sale_data3.txt']
```

</details>
<br>
<hr>

### shutil 사용법
<details>
<summary>접기/펼치기</summary>
<br>

파일 100개 이름 변경의 파일 입출력 예제에서 생성한 `일기.txt`, `일기1.txt` 파일을 코드에 적용한다.

#### 폴더 이동
```py
import shutil
shutil.move('../카스', './') # 현재 디렉토리의 상위에 있는 카스 디렉토리를 현재 디렉토리로 이동
```

#### 파일 이동
```py
import shutil
shutil.move('일기.txt', '카스')
shutil.move('일기1.txt', '카스')
```

#### 파일 복사
```py
import shutil
shutil.copy('카스/일기.txt', '테라') # 카스/일기.txt를 테라 디렉토리로 복사
```

</details>
<br>
<hr>

## 폴더 자동정리 프로그램
<details>
<summary>접기/펼치기</summary>
<br>

### 구현할 내용
1. `보고서` 단어가 포함된 파일들을 `보고서파일` 폴더로 이동
2. `보고서` 단어가 포함된 파일들을 `데이터파일` 폴더로 이동
3. `계산서` 단어가 포함된 파일들을 `계산서파일` 폴더로 이동  

### 구현 과정
1번 파일명에 `보고서` 단어가 포함된 파일들을 추출하여 `보고서파일` 폴더로 이동하는 프로그램에 대한 구현 과정을 작성한다.

#### 1. [공유폴더.zip](https://cafe.naver.com/startcodingofficial/2) 다운로드 및 공유폴더 디렉토리 생성(압축 풀기)

#### 2. `보고서` 단어가 포함된 파일 추출
glob을 활용하여 지정한 단어가 포함된 파일을 추출한다.
```py
import glob
glob.glob('C:/Users/프로젝트폴더/공유폴더/*보고서*')
```
#### 3. 폴더가 없을경우 생성
추출한 파일을 폴더로 이동시키기 전 폴더가 있는지 체크한 후 없을경우 생성한다.
```py
import os
path = 'C:/Users/프로젝트폴더/공유폴더/보고서파일'
if not os.path.exists(path):
  os.mkdir(path)
```
#### 4. 파일을 폴더로 이동
2번 과정에서 추출한 파일들을 지정한 폴더로 이동시킨다.
```py
import glob, shutil
for i in glob.glob('C:/Users/프로젝트폴더/공유폴더/*보고서*'):
  shutil.move(i, path)
```

#### 5. 데이터, 계산서 로직 구현
```py
import os, glob, shutil
path = 'C:/Users/프로젝트폴더/공유폴더/데이터파일'
if not os.path.exists(path):
  os.mkdir(path)
for i in glob.glob('C:/Users/프로젝트폴더/공유폴더/*데이터*'):
  shutil.move(i, path)
```
```py
import os, glob, shutil
path = 'C:/Users/프로젝트폴더/공유폴더/계산서파일'
if not os.path.exists(path):
  os.mkdir(path)
for i in glob.glob('C:/Users/프로젝트폴더/공유폴더/*계산서*'):
  shutil.move(i, path)
```

#### 6. 보고서, 데이터, 계산서 단어에 대한 자동화 로직 구현
```py
import os, glob, shutil
num = 1
path = f'C:/Users/~/프로젝트폴더/공유폴더_심화{num}'
shutil.copytree(r'C:\Users\프로젝트폴더\공유폴더_origin', path)

keyword_list = ['보고서', '데이터', '계산서']
for keyword in keyword_list:
  path = f'C:/Users/프로젝트폴더/공유폴더_심화{num}/{keyword}파일'
  if not os.path.exists(path):
    os.mkdir(path)
  for i in glob.glob(f'C:/Users/프로젝트폴더/공유폴더_심화{num}/*{keyword}*'):
    shutil.move(i, path)
```

**[escape raw string]**
```py
num = 2
path = rf'C:\Users\프로젝트폴더\공유폴더_심화{num}'
shutil.copytree(r'C:\Users\프로젝트폴더\공유폴더_origin', path)

keyword_list = ['보고서', '데이터', '계산서']
for keyword in keyword_list:
  path = rf'C:\Users\프로젝트폴더\공유폴더_심화{num}\{keyword}파일'
  if not os.path.exists(path):
    os.mkdir(path)
  for i in glob.glob(rf'C:\Users\프로젝트폴더\공유폴더_심화{num}\/*{keyword}*'):
    shutil.move(i, path)
```


**[강의 최종 코드]**  
```py
import os, glob, shutil
num = 4
target_folder = r'C:\Users\프로젝트폴더\공유폴더_심화'
shutil.copytree(r'C:\Users\프로젝트폴더\공유폴더_origin', f'{target_folder}{num}')

keyword_list = ['보고서', '데이터', '계산서']
for keyword in keyword_list:
  file_list = glob.glob(f'{target_folder}{num}\/*{keyword}*')
  if not os.path.exists(f'{target_folder}{num}\{keyword}파일'):
    os.mkdir(f'{target_folder}{num}\{keyword}파일')
  for i in file_list:
    shutil.move(i, f'{target_folder}{num}\{keyword}파일')
```

</details>
<br>
<hr>

### 여러 액셀파일 내용 변경
<details>
<summary>접기/펼치기</summary>
<br>

#### 실습1.
```
여러개의 액셀파일을 바꾸는 프로그램을 만든다.
`C:\Users\프로젝트폴더\공유폴더\계산서폴더` 결로에 존재하는 `계산서` 단어가 포함된 5개의 액셀파일들중
 `자료입력페이지` 시트의 `작성일자`를 `2030-02-01` 값으로 한번에 수정
```

##### 1. 한개의 액셀파일 자동화
```py
import xlwings as xw
app = xw.App(add_book=False)
wb = app.books.open(r'C:\Users\프로젝트폴더\공유폴더\계산서파일\세금계산서_놀부전자.xlsx')
ws = wb.sheets['자료입력페이지']
ws.range('C14').value = '2030-02-01'
wb.save()
app.quit()
```

##### 2. 파일목록 추출
```py
import glob
for file in glob.glob(r'C:\Users\프로젝트폴더\공유폴더\계산서파일\*.xlsx'):
  print(file)
```

##### 3. 추출된 파일 목록에 액셀 자동화 적용
```py
import xlwings as xw, glob
for file in glob.glob(r'C:\Users\프로젝트폴더\공유폴더\계산서파일\*.xlsx'):
  app = xw.App(add_book=False)
  wb = app.books.open(file)
  ws = wb.sheets['자료입력페이지']
  ws.range('C14').value = '2030-02-01'
  wb.save()
  app.quit()
```

##### 4. 최적화
```py
import xlwings as xw, glob
  app = xw.App(add_book=False)
for file in glob.glob(r'C:\Users\프로젝트폴더\공유폴더\계산서파일\*.xlsx'):
  wb = app.books.open(file)
  ws = wb.sheets['자료입력페이지']
  ws.range('C14').value = '2030-02-01'
  wb.save()
app.quit()
```

</details>
<br>
<hr>


## Template
<details>
<summary>접기/펼치기</summary>
<br>

</details>
<br>
<hr>


