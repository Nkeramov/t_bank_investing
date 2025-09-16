from tinkoff.invest.schemas import (OperationType, OperationState, TradeDirection, SecurityTradingStatus,
                                    InstrumentType, IndicatorType, TypeOfPrice)


operations_types = {
    OperationType.OPERATION_TYPE_UNSPECIFIED: 'Неизвестно',
    OperationType.OPERATION_TYPE_INPUT: 'Пополнение брокерского счета',
    OperationType.OPERATION_TYPE_OUTPUT: 'Вывод денег',
    OperationType.OPERATION_TYPE_BUY_CARD: 'Покупка с карты',
    OperationType.OPERATION_TYPE_BUY: 'Покупка',
    OperationType.OPERATION_TYPE_SELL: 'Продажа',
    OperationType.OPERATION_TYPE_BROKER_FEE: 'Комиссия брокера',
    OperationType.OPERATION_TYPE_DIVIDEND: 'Выплата дивидендов',
    OperationType.OPERATION_TYPE_TAX: 'Налоги',
    OperationType.OPERATION_TYPE_OVERNIGHT: 'Овернайт',
    OperationType.OPERATION_TYPE_DIVIDEND_TAX: 'Налоги c дивидендов',
    OperationType.OPERATION_TYPE_TAX_CORRECTION: 'Возврат налога',
    OperationType.OPERATION_TYPE_SERVICE_FEE: 'Комиссия за обслуживание',
    OperationType.OPERATION_TYPE_BENEFIT_TAX: 'Налог за материальную выгоду',
    OperationType.OPERATION_TYPE_BENEFIT_TAX_PROGRESSIVE: 'Налог за материальную выгоду по ставке 15%'
}

operations_states = {
    OperationState.OPERATION_STATE_UNSPECIFIED: 'Неизвестно',
    OperationState.OPERATION_STATE_EXECUTED: 'Выполнена',
    OperationState.OPERATION_STATE_CANCELED: 'Отменена',
    OperationState.OPERATION_STATE_PROGRESS: 'В процессе'
}

trade_directions = {
    TradeDirection.TRADE_DIRECTION_UNSPECIFIED: 'Направление сделки не определено',
    TradeDirection.TRADE_DIRECTION_BUY: 'Покупка',
    TradeDirection.TRADE_DIRECTION_SELL: 'Продажа'
}

trading_statuses = {
    SecurityTradingStatus.SECURITY_TRADING_STATUS_UNSPECIFIED: 'Торговый статус не определен',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_NOT_AVAILABLE_FOR_TRADING: 'Недоступен для торгов',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_OPENING_PERIOD: 'Период открытия торгов',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_CLOSING_PERIOD: 'Период закрытия торгов',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_BREAK_IN_TRADING: 'Перерыв в торговле',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_NORMAL_TRADING: 'Нормальная торговля',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_CLOSING_AUCTION: 'Аукцион закрытия',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_DARK_POOL_AUCTION: 'Аукцион крупных пакетов',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_DISCRETE_AUCTION: 'Дискретный аукцион',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_OPENING_AUCTION_PERIOD: 'Аукцион открытия',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_TRADING_AT_CLOSING_AUCTION_PRICE: 'Период торгов по цене аукциона закрытия',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_SESSION_ASSIGNED: 'Сессия назначена',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_SESSION_CLOSE: 'Сессия закрыта',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_SESSION_OPEN: 'Сессия открыта',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_DEALER_NORMAL_TRADING: 'Доступна торговля в режиме внутренней ликвидности брокера',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_DEALER_BREAK_IN_TRADING: 'Перерыв торговли в режиме внутренней ликвидности брокера',
    SecurityTradingStatus.SECURITY_TRADING_STATUS_DEALER_NOT_AVAILABLE_FOR_TRADING: 'Недоступна торговля в режиме внутренней ликвидности брокера'
}

instrument_types = {
    InstrumentType.INSTRUMENT_TYPE_UNSPECIFIED: 'Тип инструмента не определен',
    InstrumentType.INSTRUMENT_TYPE_BOND: 'Облигация',
    InstrumentType.INSTRUMENT_TYPE_SHARE: 'Акция',
    InstrumentType.INSTRUMENT_TYPE_CURRENCY: 'Валюта',
    InstrumentType.INSTRUMENT_TYPE_ETF: 'Фонд',
    InstrumentType.INSTRUMENT_TYPE_FUTURES: 'Фьючерс',
    InstrumentType.INSTRUMENT_TYPE_SP: 'Структурная нота',
    InstrumentType.INSTRUMENT_TYPE_OPTION: 'Опцион',
    InstrumentType.INSTRUMENT_TYPE_CLEARING_CERTIFICATE: 'Клиринговый сертификат участия',
    InstrumentType.INSTRUMENT_TYPE_INDEX: 'Индекс',
    InstrumentType.INSTRUMENT_TYPE_COMMODITY: 'Товар'
}

price_types = {
    TypeOfPrice.TYPE_OF_PRICE_UNSPECIFIED: 'Не указано',
    TypeOfPrice.TYPE_OF_PRICE_CLOSE: 'Цена закрытия',
    TypeOfPrice.TYPE_OF_PRICE_OPEN: 'Цена открытия',
    TypeOfPrice.TYPE_OF_PRICE_HIGH: 'Максимальное значение за выбранный интервал',
    TypeOfPrice.TYPE_OF_PRICE_LOW: 'Минимальное значение за выбранный интервал',
    TypeOfPrice.TYPE_OF_PRICE_AVG: 'Среднее значение по показателям [ (close + open + high + low) / 4 ]'
}

indicator_types = {
    IndicatorType.INDICATOR_TYPE_UNSPECIFIED: 'Не определен',
    IndicatorType.INDICATOR_TYPE_BB: 'Линия Боллинжера',
    IndicatorType.INDICATOR_TYPE_EMA: 'Экспоненциальная скользящая средняя',
    IndicatorType.INDICATOR_TYPE_RSI: 'Индекс относительной силы',
    IndicatorType.INDICATOR_TYPE_MACD: 'Схождение/расхождение скользящих средних',
    IndicatorType.INDICATOR_TYPE_SMA: 'Простое скользящее среднее'
}