import http.server
import socketserver
import urllib.parse
import os

PORT = 8080
DIRECTORY = "/Users/jeonhyochul/work/Websquare"

class WebSquareHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/download_excel':
            params = urllib.parse.parse_qs(parsed.query)
            fmt = params.get('format', ['xls'])[0]
            
            if fmt == 'xlsx':
                filename = "WebSquare_Report.xlsx"
                filepath = os.path.join(DIRECTORY, filename)
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                filename = "WebSquare_Report.xls"
                filepath = os.path.join(DIRECTORY, filename)
                content_type = "application/vnd.ms-excel; charset=utf-8"

            if os.path.exists(filepath):
                with open(filepath, "rb") as f:
                    excel_data = f.read()
            else:
                excel_data = self.generate_excel_content(fmt)

            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(excel_data)))
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.end_headers()
            self.wfile.write(excel_data)
            return

        return super().do_GET()

    def generate_excel_content(self, fmt):
        xml = '''<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--[if gte mso 9]><xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>
  <x:ExcelWorksheet><x:Name>공급실 작업 관리</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>설비 정기점검</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>교체 자재 재고</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
</x:ExcelWorksheets>
</x:ExcelWorkbook>
</xml><![endif]-->
<style>
  body { font-family: "Malgun Gothic", "Pretendard", sans-serif; }
  .ws-title { font-size: 15pt; font-weight: bold; color: #ffffff; background-color: #1e293b; text-align: left; padding: 10px 14px; height: 38px; border: 1px solid #0f172a; }
  .ws-info { font-size: 9.5pt; color: #475569; background-color: #f1f5f9; padding: 6px; height: 24px; border: 1px solid #cbd5e1; }
  .ws-th { background-color: #2563eb; color: #ffffff; font-weight: bold; text-align: center; border: 1px solid #1d4ed8; padding: 9px; font-size: 11pt; height: 28px; }
  .ws-td { border: 1px solid #cbd5e1; padding: 7px 10px; font-size: 10pt; color: #1e293b; vertical-align: middle; }
  .ws-td-even { background-color: #f8fafc; }
  .ws-seq { text-align: center; font-weight: bold; background-color: #e2e8f0; color: #334155; width: 50px; }
  .ws-center { text-align: center; }
  .ws-right { text-align: right; }
  .ws-sum { background-color: #e0e7ff; font-weight: bold; color: #3730a3; border: 1.5px solid #4f46e5; height: 32px; font-size: 10.5pt; }
</style>
</head>
<body>
<table style="width:100%; border-collapse:collapse; margin-bottom:30px;">
<tr><td colspan="6" class="ws-title">📊 [WebSquare Grid Report] 1. 공급실 작업 관리 현황</td></tr>
<tr><td colspan="6" class="ws-info">출력 시스템: WebSquare Grid Engine | 리포트 구분: 3개 그리드 멀티시트 통합 리포트</td></tr>
<tr><td colspan="6" style="height:10px;"></td></tr>
<tr>
  <th class="ws-th">NO</th>
  <th class="ws-th">작업일자</th>
  <th class="ws-th">공급실명</th>
  <th class="ws-th">담당자</th>
  <th class="ws-th">소독상태</th>
  <th class="ws-th">특이사항</th>
</tr>
<tr><td class="ws-td ws-seq">1</td><td class="ws-td ws-center">2026-09-01</td><td class="ws-td" style="font-weight:bold;">클린룸 A동</td><td class="ws-td">김철수</td><td class="ws-td ws-center">양호</td><td class="ws-td">정기 소독 완료</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">2</td><td class="ws-td ws-center ws-td-even">2026-09-02</td><td class="ws-td ws-td-even" style="font-weight:bold;">무균 준비실</td><td class="ws-td ws-td-even">정수진</td><td class="ws-td ws-center ws-td-even">양호</td><td class="ws-td ws-td-even">필터 점검 병행</td></tr>
<tr><td class="ws-td ws-seq">3</td><td class="ws-td ws-center">2026-09-03</td><td class="ws-td" style="font-weight:bold;">원자재 보관실</td><td class="ws-td">한지은</td><td class="ws-td ws-center">주의</td><td class="ws-td">습도 조절 완료</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">4</td><td class="ws-td ws-center ws-td-even">2026-09-04</td><td class="ws-td ws-td-even" style="font-weight:bold;">포장 작업실</td><td class="ws-td ws-td-even">유재석</td><td class="ws-td ws-center ws-td-even">양호</td><td class="ws-td ws-td-even">특이사항 없음</td></tr>
<tr><td class="ws-td ws-seq">5</td><td class="ws-td ws-center">2026-09-05</td><td class="ws-td" style="font-weight:bold;">클린룸 B동</td><td class="ws-td">이영희</td><td class="ws-td ws-center">양호</td><td class="ws-td">UV 램프 교체</td></tr>
<tr><td colspan="6" class="ws-sum" style="padding:10px; text-align:right;">총 작업 항목 수: 5건 (정상 완료)</td></tr>
</table>

<br style="page-break-before:always;">

<table style="width:100%; border-collapse:collapse; margin-bottom:30px;">
<tr><td colspan="6" class="ws-title">📊 [WebSquare Grid Report] 2. 설비 정기점검 현황</td></tr>
<tr><td colspan="6" class="ws-info">출력 시스템: WebSquare Grid Engine</td></tr>
<tr><td colspan="6" style="height:10px;"></td></tr>
<tr>
  <th class="ws-th">NO</th>
  <th class="ws-th">점검일자</th>
  <th class="ws-th">설비명</th>
  <th class="ws-th">점검항목</th>
  <th class="ws-th">점검결과</th>
  <th class="ws-th">점검자</th>
</tr>
<tr><td class="ws-td ws-seq">1</td><td class="ws-td ws-center">2026-09-01</td><td class="ws-td" style="font-weight:bold;">공조기 (AHU-01)</td><td class="ws-td">풍량 및 필터 차압 검사</td><td class="ws-td ws-center">합격</td><td class="ws-td">박민수</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">2</td><td class="ws-td ws-center ws-td-even">2026-09-03</td><td class="ws-td ws-td-even" style="font-weight:bold;">순수제조장치 (RO/EDI)</td><td class="ws-td ws-td-even">비전도도 측정</td><td class="ws-td ws-center ws-td-even">합격</td><td class="ws-td ws-td-even">윤서준</td></tr>
<tr><td class="ws-td ws-seq">3</td><td class="ws-td ws-center">2026-09-05</td><td class="ws-td" style="font-weight:bold;">냉동기 (Chiller-02)</td><td class="ws-td">냉매 압력 및 누설 점검</td><td class="ws-td ws-center">보수필요</td><td class="ws-td">송지효</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">4</td><td class="ws-td ws-center ws-td-even">2026-09-07</td><td class="ws-td ws-td-even" style="font-weight:bold;">보일러 (Boiler-01)</td><td class="ws-td ws-td-even">안전밸브 작동 점검</td><td class="ws-td ws-center ws-td-even">합격</td><td class="ws-td ws-td-even">김종국</td></tr>
<tr><td class="ws-td ws-seq">5</td><td class="ws-td ws-center">2026-09-09</td><td class="ws-td" style="font-weight:bold;">배기팬 (Exhaust Fan)</td><td class="ws-td">진동 및 소음 측정</td><td class="ws-td ws-center">합격</td><td class="ws-td">최동현</td></tr>
<tr><td colspan="6" class="ws-sum" style="padding:10px; text-align:right;">총 점검 건수: 5건 | 합격: 4건 | 보수필요: 1건</td></tr>
</table>

<br style="page-break-before:always;">

<table style="width:100%; border-collapse:collapse; margin-bottom:30px;">
<tr><td colspan="6" class="ws-title">📊 [WebSquare Grid Report] 3. 교체 자재 재고 현황</td></tr>
<tr><td colspan="6" class="ws-info">출력 시스템: WebSquare Grid Engine</td></tr>
<tr><td colspan="6" style="height:10px;"></td></tr>
<tr>
  <th class="ws-th">NO</th>
  <th class="ws-th">입고일자</th>
  <th class="ws-th">자재코드</th>
  <th class="ws-th">자재명</th>
  <th class="ws-th">재고수량</th>
  <th class="ws-th">단가 (원)</th>
</tr>
<tr><td class="ws-td ws-seq">1</td><td class="ws-td ws-center">2026-08-20</td><td class="ws-td ws-center" style="font-weight:bold; color:#2563eb;">MAT-H14-01</td><td class="ws-td" style="font-weight:bold;">HEPA 필터 (H14 등급)</td><td class="ws-td ws-right" style="font-weight:bold;">45 개</td><td class="ws-td ws-right">₩ 120,000</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">2</td><td class="ws-td ws-center ws-td-even">2026-08-25</td><td class="ws-td ws-center ws-td-even" style="font-weight:bold; color:#2563eb;">MAT-UV-02</td><td class="ws-td ws-td-even" style="font-weight:bold;">UV 살균 램프 40W</td><td class="ws-td ws-right ws-td-even" style="font-weight:bold;">120 개</td><td class="ws-td ws-right ws-td-even">₩ 35,000</td></tr>
<tr><td class="ws-td ws-seq">3</td><td class="ws-td ws-center">2026-08-28</td><td class="ws-td ws-center" style="font-weight:bold; color:#2563eb;">MAT-RO-03</td><td class="ws-td" style="font-weight:bold;">RO 멤브레인 필터 8인치</td><td class="ws-td ws-right" style="font-weight:bold;">15 개</td><td class="ws-td ws-right">₩ 480,000</td></tr>
<tr><td class="ws-td ws-seq ws-td-even">4</td><td class="ws-td ws-center ws-td-even">2026-09-02</td><td class="ws-td ws-center ws-td-even" style="font-weight:bold; color:#2563eb;">MAT-CH-04</td><td class="ws-td ws-td-even" style="font-weight:bold;">냉동기 컴프레서 오일 20L</td><td class="ws-td ws-right ws-td-even" style="font-weight:bold;">8 개</td><td class="ws-td ws-right ws-td-even">₩ 210,000</td></tr>
<tr><td class="ws-td ws-seq">5</td><td class="ws-td ws-center">2026-09-05</td><td class="ws-td ws-center" style="font-weight:bold; color:#2563eb;">MAT-BL-05</td><td class="ws-td" style="font-weight:bold;">보일러 가스켓 개스킷 세트</td><td class="ws-td ws-right" style="font-weight:bold;">60 개</td><td class="ws-td ws-right">₩ 18,000</td></tr>
<tr><td colspan="6" class="ws-sum" style="padding:10px; text-align:right;">총 자재 품목: 5종류 | 총 재고수량: 248개 | 총 재고금액 합계: ₩ 18,560,000</td></tr>
</table>
</body></html>'''
        return ('\ufeff' + xml).encode('utf-8')

if __name__ == '__main__':
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), WebSquareHandler) as httpd:
        print(f"Serving HTTP on port {PORT} with Excel Download API...")
        httpd.serve_forever()
