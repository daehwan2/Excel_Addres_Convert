# Excel_Addres_Convert
 
C# Windows Forms로 만든 엑셀 파일의 지번 주소를 도로명 주소로 변환하는 프로그램입니다.

Selenium 과 C# 엑셀 라이브러리를 사용하였습니다.

# 기획 배경

2020년에 공익 근무를 하던 중 지번주소(구주소)를 도로명주소(신주소)로 바꾸는 일을 했었습니다.

엑셀파일로 작성되어있었는데 주소갯수가 500개는 됐었습니다.

하나하나 검색해서 일을 했었는데 단순 반복 작업이였고 프로그램으로 만들면 1분 걸릴걸 3시간을 넘게 작업을 했었습니다. 

그래서 퇴근을 하고 집가서 바로 당시 공부하고있던 C#으로 작업을 진행하였었습니다.

근무하는 곳이 보안이 철저한 곳이라 아무프로그램이나 다운못받게하여서 실제로 작업에 쓰이진 못했습니다.


# 요구사항 명세

- 지번 주소가 나열되어 있는 엑셀파일을 선택한다.
- 시트를 선택한뒤 지번주소의 열과 변환해서 도로명주소를 넣을 열을 선택한다.
- 시작행과 끝행을 지정한다.
- 변환을 누르면 http://www.juso.go.kr/openIndexPage.do 페이지로 이동해서 자동으로 검색한뒤 엑셀파일에 붙여준다.


# 결과물 

https://blog.naver.com/dhdh0482/221883588116

당시에 운영하던 블로그에 배포를 하였고 사용법은 동영상 형태로 찍어서 제공하였다.

아무나 사용할 수있게 끔 하였는데 생각보다 많은 사람이 이용해주어서 뿌듯하였다.

![image](https://user-images.githubusercontent.com/53414542/162556129-7e14b6d9-5f04-408b-a0e1-a1663f20f504.png)
