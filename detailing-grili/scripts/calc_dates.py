#!/usr/bin/env python3
"""Номер пропозиції (ДДММРР) і термін дії (+3 робочі дні, сб і нд не рахуються).
Використання: python3 calc_dates.py ДД.ММ.РРРР   (без аргументу — сьогодні)
Державні свята НЕ враховуються: якщо термін потрапляє на свято — повідомте менеджера."""
import sys
from datetime import date, datetime, timedelta

d = datetime.strptime(sys.argv[1], "%d.%m.%Y").date() if len(sys.argv) > 1 else date.today()
v, n = d, 0
while n < 3:
    v += timedelta(days=1)
    if v.weekday() < 5:
        n += 1
wd = ["понеділок", "вівторок", "середа", "четвер", "пʼятниця", "субота", "неділя"]
print(f"дата подачі:  {d:%d.%m.%Y} ({wd[d.weekday()]})")
print(f"номер:        {d:%d%m%y}   (якщо вже є за цей день — додайте _1, _2)")
print(f"діє до:       {v:%d.%m.%Y} ({wd[v.weekday()]})")
