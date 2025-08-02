from tinkoff.invest import Client, PortfolioResponse, Quotation, MoneyValue, Operation, OperationState, InstrumentIdType
from tinkoff.invest.utils import quotation_to_decimal
import pandas as pd
from pathlib import Path
import sys
from datetime import datetime, timedelta, timezone
from decimal import Decimal, getcontext

import api_token

# Настройка точности вычислений
getcontext().prec = 10


def convert_tinkoff_value(value):
    """Конвертирует Quotation или MoneyValue в float"""
    if isinstance(value, (Quotation, MoneyValue)):
        units = Decimal(str(value.units))
        nano = Decimal(str(value.nano))
        return float(units + nano / Decimal('1e9'))
    return value


def get_portfolio_data(client, account_id):
    """Получает данные портфеля"""
    portfolio = client.operations.get_portfolio(account_id=account_id)

    all_fields = [
        'figi', 'instrument_type', 'quantity',
        'average_position_price', 'expected_yield',
        'current_nkd', 'average_position_price_pt',
        'current_price', 'average_position_price_fifo',
        'quantity_lots', 'blocked', 'blocked_lots',
        'position_uid', 'instrument_uid', 'var_margin',
        'expected_yield_fifo', 'daily_yield', 'ticker'
    ]

    positions = []
    for position in portfolio.positions:
        pos_data = {}
        for field in all_fields:
            value = getattr(position, field, None)
            if value is not None:
                value = convert_tinkoff_value(value)
            pos_data[field] = value if value is not None else (0.0 if field in [
                'current_nkd', 'var_margin', 'daily_yield'] else '')
        positions.append(pos_data)

    return pd.DataFrame(positions).rename(columns={
        'figi': 'FIGI',
        'instrument_type': 'Тип',
        'quantity': 'Количество',
        'average_position_price': 'Средняя цена',
        'expected_yield': 'Доходность',
        'current_nkd': 'НКД',
        'average_position_price_pt': 'Ср.цена (пункты)',
        'current_price': 'Текущая цена',
        'average_position_price_fifo': 'Ср.цена FIFO',
        'quantity_lots': 'Кол-во лотов',
        'blocked': 'Заблокировано',
        'blocked_lots': 'Заблок. лотов',
        'position_uid': 'UID позиции',
        'instrument_uid': 'UID инструмента',
        'var_margin': 'Var Margin',
        'expected_yield_fifo': 'Доход.FIFO',
        'daily_yield': 'Дневной доход',
        'ticker': 'Тикер',
    })


def get_operations_history(client, account_id, days=3650):
    """Получает историю операций за указанный период с детализацией"""
    from tinkoff.invest import OperationState, OperationType

    to_date = datetime.now(timezone.utc)
    from_date = to_date - timedelta(days=days)

    operations = client.operations.get_operations(
        account_id=account_id,
        from_=from_date,
        to=to_date,
        state=OperationState.OPERATION_STATE_EXECUTED
    ).operations

    instrument_cache = {}

    operations_data = []
    for op in operations:
        if op.figi not in instrument_cache:
            try:
                instrument = client.instruments.get_instrument_by(
                    id_type=InstrumentIdType.INSTRUMENT_ID_TYPE_FIGI,
                    id=op.figi
                ).instrument
                instrument_cache[op.figi] = {
                    'ticker': instrument.ticker,
                    'name': instrument.name
                }
            except:
                instrument_cache[op.figi] = {
                    'ticker': '',
                    'name': ''
                }

        instrument_info = instrument_cache[op.figi]

        operation_type = {
            OperationType.OPERATION_TYPE_BUY: "Покупка",
            OperationType.OPERATION_TYPE_SELL: "Продажа",
            OperationType.OPERATION_TYPE_DIVIDEND: "Дивиденды",
            OperationType.OPERATION_TYPE_DIVIDEND_TAX: "Налог на дивиденды",
            OperationType.OPERATION_TYPE_BROKER_FEE: "Комиссия брокера",
            OperationType.OPERATION_TYPE_SERVICE_FEE: "Комиссия за обслуживание",
        }.get(op.operation_type, op.type)

        op_data = {
            'ID операции': op.id,
            'Дата': op.date.replace(tzinfo=None),
            'Тип операции': operation_type,
            'FIGI': op.figi,
            'Тикер': instrument_info['ticker'],
            'Название': instrument_info['name'],
            'Количество': convert_tinkoff_value(op.quantity) if op.quantity else 0,
            'Цена за единицу': convert_tinkoff_value(op.price) if op.price else 0,
            'Сумма операции': convert_tinkoff_value(op.payment),
            'Валюта': op.currency,
            'Статус': 'Исполнена',
            'Комиссия': convert_tinkoff_value(getattr(op, 'commission', 0)) or 0,
            'Тип инструмента': op.instrument_type,
        }

        if hasattr(op, 'parent_operation_id') and op.parent_operation_id:
            op_data['Родительская операция'] = op.parent_operation_id

        operations_data.append(op_data)

    df = pd.DataFrame(operations_data)

    if not df.empty:
        df = df.sort_values('Дата', ascending=False)
        df['Дата'] = pd.to_datetime(df['Дата']).dt.strftime('%d.%m.%Y %H:%M')

    return df


def save_to_excel(api_token: str, output_file: str = "invest_report.xlsx", account_id: str = None):
    """Сохраняет портфель и историю операций в Excel с форматированием"""
    try:
        import xlsxwriter
    except ImportError:
        print("Ошибка: Требуется xlsxwriter. Установите: pip install xlsxwriter", file=sys.stderr)
        return

    with Client(api_token) as client:
        if not account_id:
            account_id = client.users.get_accounts().accounts[0].id

        portfolio_df = get_portfolio_data(client, account_id)
        operations_df = get_operations_history(client, account_id)

        output_path = Path(output_file).with_suffix('.xlsx')
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            portfolio_df.to_excel(writer, sheet_name='Портфель', index=False)
            operations_df.to_excel(writer, sheet_name='Операции', index=False)

            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'dd.mm.yyyy hh:mm'})
            header_fmt = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })

            # Форматируем лист Портфель
            portfolio_sheet = writer.sheets['Портфель']
            for col_num, value in enumerate(portfolio_df.columns.values):
                portfolio_sheet.write(0, col_num, value, header_fmt)
                if value in ['Средняя цена', 'Текущая цена', 'Доходность', 'НКД']:
                    portfolio_sheet.set_column(col_num, col_num, 15, money_fmt)

            # Форматируем лист Операции
            operations_sheet = writer.sheets['Операции']
            for col_num, value in enumerate(operations_df.columns.values):
                operations_sheet.write(0, col_num, value, header_fmt)
                if value == 'Дата':
                    operations_sheet.set_column(col_num, col_num, 18, date_fmt)
                elif value in ['Цена за единицу', 'Сумма операции', 'Комиссия']:
                    operations_sheet.set_column(col_num, col_num, 15, money_fmt)

            # Условное форматирование по типу операции
            green_format = workbook.add_format({'bg_color': '#C6EFCE'})
            red_format = workbook.add_format({'bg_color': '#FFC7CE'})

            operations_sheet.conditional_format(
                'C2:C1000', {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'Дивиденды',
                    'format': green_format
                }
            )
            operations_sheet.conditional_format(
                'C2:C1000', {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'Налог',
                    'format': red_format
                }
            )

        print(f"Отчёт сохранён: {output_path.resolve()}")


if __name__ == "__main__":
    TOKEN = api_token.TOKEN  # Убедитесь, что файл api_token.py содержит строку: TOKEN = "ваш_токен"
    save_to_excel(TOKEN, "investment_report.xlsx")
