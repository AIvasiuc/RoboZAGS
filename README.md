# RoboZAGS
Данный модуль представляет собой финальный этап регистрации брака. Исходя из всех ранее полученных данных (ФИО,
ответы на вопросы для определения совместимости пары, электронная потча, фото, присвоенный статус отношений) формируется сертификат - 
лист, который будет отправлен на принтер на печать. Здесь же есть элементы работы с COM портом (для микроконтроллера Arduino, который
отправляет сигнал двум двигателям, работа которых заключается в выдаче 2-х колец посетителям). В итоге посетители получают на руки сертификат
и кольца + сформированный сертификат отправляется им на почту(которую они указывают на одной из страниц). 
Готовый продукт представляет собой интерактивное WPF приложение. Программа работает с периферийными устройствами (принтер,
веб - камера и микроконтроллер Arduino).
