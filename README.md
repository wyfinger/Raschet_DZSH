Макрос ДЗШ
==========

Макрос предназначен для службы расчета параметров настроек устройств релейной защиты и автоматики Приморского РДУ.

Задача макроса - подготовка "карты токов короткого замыкания" для последующего расчета уставок дифференциальных защит шин.

Макрос предназначен для использования вместе с ПК АРМ СРЗА (модуль ТКЗ-2000) и в среде Microsoft Excel.

Результаты работы макроса и, соответственно, этапы работы c ним делятся на несколько этапов:

1. Подготовка исходных данных для макроса - экспорт схемы сети, для которой должен производиться расчет, в формат Microsoft Excel. В любой момент времени мы используем две сети: с максимальным составом оборудования и с минимальным. Для расчета ДЗШ главным образом требуется минимальная схема. Данную схему необходимо конвертировать в формат Excel.

2. До запуска макроса необходимо удостовериться в том, что программа ПК АРМ СРЗА (модуль ТКЗ-2000) запущена и в нее загружена именно та схема, которая была конвертирована в Excel. При этом режим работы программы - приказы или диалоговый - не имеет значения, макрос будет манипулировать программой ТКЗ-2000 самостоятельно.

3. При запуске макроса потребуется указать узел для расчета. На основе данных сети макрос подготовит два приказа, скопирует их в ТКЗ-2000, выполнит расчет, проанализирует протоколы расчета, предложит сохранить протоколы (макрос добавляет в начало протокола расчета исходный приказ с комментариями, т.к. все программы АРМ СРЗА удаляют комментарии из протокола расчета) и подготовит результирующие данные в виде нового листа Excel документа.

Результатом работы макроса являются токи короткого замыкания (3, 2, 1+1, и 1- фазного) в указанном узле в следующих режимах:

1. Нормальный режим;
2. Отключение (N-1) по одному присоединению шин (узел расчета). При этом макрос отключает и тупики, что проявляется в одинаковых замерах тока в режимах "ВСЕ ВКЛЮЧЕНО" и "N-1". Результаты данного замера используются, как правило, для проверки чувствительности выбранных уставок ДЗШ.
3. Отключение всех присоединений узла расчета кроме одного, ведущего к "питающему узлу" (режим опробования).
4. Аналогично предыдущему пункту, плюс отключение (N-1) по одному присоединению питающего узла. Результаты данного замера используются, как правило, для проверки чувствительности выбранных уставок ДЗШ в режиме опробования (чувствительный орган ДЗШ или очувствление).

Особенности и ограничения
=========================

1. Узлы сети должны именоваться только числами, работа по реализации возможности работы с буквенно-символьными наименованиями узлов в планах.
2. Ветви, трансформаторы и прочие элементы сети должны иметь номер элемента.


Разработчик - ведущий специалист СРЗА Приморского РДУ

Матвеев И.В. - [miv@prim.so-ups.ru](mailto:miv@prim.so-ups.ru)
