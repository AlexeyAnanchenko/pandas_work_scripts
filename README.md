# Автоматизации бизнес-процессов с помощью Python (Pandas, NumPy, Selenium)

Данный проект написан с целью сокращения своих трудозатрат на ручные операции, повышение качества выполняемой работы и приобретания опыта работы с популярными библиотеками обработки данных (Pandas, NumPy).

Проект демонстрационный, отображает мои навыки работы с вышеуказанными библиотеками.

Ниже краткое описание основных результатов:
1. Атоматическая регистрация заказов клиентов и изменений по ним, по датам и кол-ву. (<code>[registry_factors.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/base_scripts/registry_factors.py)
</code>)

2. Автовыгрузка SRS и PBI отчётности с помощью библиотеки Selenium (<code>[update_data.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/base_scripts/update_data.py)
</code>)

3. Автопроверка заказов клиентов на корректность (<code>[check_factors.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/reports/check_factors.py)
</code>)

4. Фиксация недогруженных заказов клиентам и распределение остатка по ответственным. (<code>[not_sold.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/reports/not_sold.py)
</code>)

5. Отчёт аккумулирующий всю возможную информация в разрезе Клиент-Склад-Штрихкод (<code>[registry_potential_sales.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/reports/registry_potential_sales.py)
</code>)

6. Eжедневное архивирование сформированных отчётов (<code>[archive.py](https://github.com/AlexeyAnanchenko/pandas_work_scripts/blob/main/scripts/base_scripts/archive.py)
</code>)