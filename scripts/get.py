import requests

from hidden_settings import user_password, data


url = 'https://r.alidi.ru/ReportServer/Pages/ReportViewer.aspx?%2f%u041b%u043e%u0433%u0438%u0441%u0442%u0438%u043a%u0430%2f%u0417%u0430%u043a%u0443%u043f%u043a%u0438%2f1082+-+%u0414%u043e%u0441%u0442%u0443%u043f%u043d%u043e%u0441%u0442%u044c+%u0442%u043e%u0432%u0430%u0440%u0430+%u043f%u043e+%u0441%u043a%u043b%u0430%u0434%u0430%u043c+(PG)&rc%3ashowbackbutton=true'
referer = 'https://r.alidi.ru/ReportServer/Pages/ReportViewer.aspx?%2F%D0%9B%D0%BE%D0%B3%D0%B8%D1%81%D1%82%D0%B8%D0%BA%D0%B0%2F%D0%97%D0%B0%D0%BA%D1%83%D0%BF%D0%BA%D0%B8%2F1082%20-%20%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BD%D0%BE%D1%81%D1%82%D1%8C%20%D1%82%D0%BE%D0%B2%D0%B0%D1%80%D0%B0%20%D0%BF%D0%BE%20%D1%81%D0%BA%D0%BB%D0%B0%D0%B4%D0%B0%D0%BC%20(PG)&rc:showbackbutton=true'

user_agent = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
              'AppleWebKit/537.36 (KHTML, like Gecko) '
              'Chrome/108.0.0.0 Safari/537.36')

session = requests.Session()
session.auth = user_password
session.headers.update({
    'User-Agent': user_agent,
    'Host': 'r.alidi.ru',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
    'Cache-Control': 'no-cache',
    'Content-Length': '35309',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://r.alidi.ru',
    'Referer': referer,
    'Sec-Fetch-Dest': 
})


Sec-Fetch-Dest: empty
Sec-Fetch-Mode: cors
Sec-Fetch-Site: same-origin
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36
X-MicrosoftAjax: Delta=true
X-Requested-With: XMLHttpRequest
sec-ch-ua: "Not?A_Brand";v="8", "Chromium";v="108", "Google Chrome";v="108"
sec-ch-ua-mobile: ?0
sec-ch-ua-platform: "Windows"


# response = session.get(url_rem)
# with open("SRS.xlsx", "wb") as file:
#     file.write(response.content)
#     file.close()


response = session.post(url, data)
print(dict(response.request.headers))
with open("SRS.text", "w") as file:
    file.write(response.request.headers)
