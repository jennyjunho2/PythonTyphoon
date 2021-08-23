##사용하는 방법

본 파일은 파이썬을 이용하여 자동으로 검토용파일을 제작하고, 수정된 검토용파일을 바탕으로 반영 및 출처 표시를 합니다.
(추후 업데이트 예정)

모르겠으면 010-3220-8149 또는 a01032208149@gmail.com으로 연락주세요.

- 작업은 공유폴더상의 파일이 아닌, 복사된 파일을 사용해주세요. 인터넷 문제로 인하여 오류를 발생시킬 수 있습니다.
- 더 쉽게 할 수 있도록 프로그램(GUI) 제작하겠습니다 ^ㅇ^

#1. 엑셀 파일을 통해 검토용파일 제작하기
엑셀 파일을 통해 검토용 파일 제작에 사용되는 함수는 source_to_basefile_execute 함수입니다.

Ex) source_to_basefile_execute(hwp = hwp, excel = excel_example, grade_number = 1, test_name = "고2 양정...)

##1.1) hwp = hwp
이 부분은 건들이지 말아주세요.

##1.2) excel = excel_example
excel =  이후에는 엑셀파일의 경로를 입력해주면 됩니다.

Ex) excel = r"C:\Users\Season\Desktop\내신주문서.xlsx"

경로를 쉽게 따오는 방법은 해당 파일을 누르고 (Shift+F10)을 누르고 A를 누르면 경로가 자동으로 복사됩니다. 이후 excel = 옆에 넣어주면 되는데, 앞의 큰따옴표 앞에 'r'를 반드시 붙여주세요. 안붙이면 백슬래시가 인식되지 않습니다.

만약 엑셀파일을 내신주문서 대신 해당 부분만 따로 떼서 하실 경우 "고1 2021" 또는 "고2 2021" 로 시트 이름을 바꾸어 주세요.


##1.3) grade_number = grade_number

고1 검토용파일 제작이면 grade_number = 1, 고2 검토용파일 제작이면 grade_number = 2를 넣어주시면 됩니다.

##1.4) test_name = "고2 양정..."

test_name = 이후에는 엑셀에서 제작할 시험지 이름을 쓰면 됩니다. 엑셀파일에서 제목을 복사-붙여넣기하면 됩니다.
여기서도 큰따옴표를 양쪽에 넣어주어야 하고, 양 옆에는 어떤 문자도 붙여주시지 않으면 됩니다. 

Ex) test_name = "1번 도형의 이동~집합과 명제 (210828)"

#2. 검토용 파일 수정 이후, 반영 및 출처표시하기

검토용파일 수정이후 반영과 출처표시를 하는 함수는 basefile_to_source 함수입니다.

Ex) basefile_to_source(hwp = hwp, basefile=basefile_example, grade_number = grade_number, excel = excel_test_2

##2.1) hwp = hwp

이 부분은 건들이지 말아주세요.

##2.2) basefile = basefile_example

basefile = 옆에는 검토용파일 경로를 입력해주면 됩니다. 경로는 엑셀파일 경로 복사하듯이 마찬가지로 (Shift+F10)과 A를 누르면 빠르게 복사할 수 있습니다. \
이후 basefile = 옆에 넣어주면 되는데, 앞의 큰따옴표 앞에 'r'를 반드시 붙여주세요. 안붙이면 백슬래시가 인식되지 않습니다. 

Ex) basefile = r"C:\Users\Season\Desktop\"C:\Users\Season\Desktop\시험용파일.hwp"

##2.3) grade_number = grade_number

고1 검토용파일이면 grade_number = 1, 고2 검토용파일이면 grade_number = 2를 넣어주시면 됩니다.

##2.4) excel = 

excel =  이후에는 엑셀파일의 경로를 입력해주면 됩니다.

Ex) excel = r"C:\Users\Season\Desktop\내신주문서.xlsx"

경로를 쉽게 따오는 방법은 해당 파일을 누르고 (Shift+F10)을 누르고 A를 누르면 경로가 자동으로 복사됩니다. 이후 excel = 옆에 넣어주면 되는데, 앞의 큰따옴표 앞에 'r'를 반드시 붙여주세요. 안붙이면 백슬래시가 인식되지 않습니다.

만약 엑셀파일을 내신주문서 대신 해당 부분만 따로 떼서 하실 경우 "고1 2021" 또는 "고2 2021" 로 시트 이름을 바꾸어 주세요.
