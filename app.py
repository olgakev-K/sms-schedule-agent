def extract_milestones_from_sheet(wb):
    """
    Надежно извлекает вехи из вкладки 'START PROJECT TOGF-ENG-007-02'
    Возвращает список словарей: [{"name": "...", "date": "...", "week": ...}, ...]
    """
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    milestones = []
    
    if sheet_name not in wb.sheetnames:
        return milestones
    
    ws = wb[sheet_name]
    
    # Шаг 1: Находим колонку с названиями вех
    milestone_col = None
    date_col = None
    
    for row in range(1, 10):  # Ищем заголовки в первых 10 строках
        for col in range(1, 20):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                val_lower = str(cell.value).lower()
                
                # Ищем колонку с вехами
                if any(kw in val_lower for kw in ["milestone", "веха", "этап", "gate", "review", "name", "наименование"]):
                    milestone_col = col
                
                # Ищем колонку с датами
                if any(kw in val_lower for kw in ["date", "дата", "deadline", "срок", "target"]):
                    date_col = col
    
    # Если не нашли по заголовкам, используем эвристику
    if milestone_col is None:
        milestone_col = 1  # По умолчанию первая колонка
    if date_col is None:
        date_col = milestone_col + 1  # Дата скорее всего в следующей колонке
    
    # Шаг 2: Извлекаем данные вех
    for row in range(2, min(ws.max_row + 1, 100)):
        milestone_name = ws.cell(row=row, column=milestone_col).value
        milestone_date = ws.cell(row=row, column=date_col).value
        
        if milestone_name and str(milestone_name).strip():
            # Очищаем название
            name_clean = str(milestone_name).strip()
            
            # Пытаемся распарсить дату
            date_parsed = None
            if isinstance(milestone_date, datetime):
                date_parsed = milestone_date.date()
            elif milestone_date and isinstance(milestone_date, str):
                try:
                    date_parsed = datetime.strptime(milestone_date, "%d.%m.%Y").date()
                except:
                    try:
                        date_parsed = datetime.strptime(milestone_date, "%Y-%m-%d").date()
                    except:
                        pass
            
            milestones.append({
                "name": name_clean,
                "date": date_parsed,
                "row": row
            })
    
    return milestones
