import pandas as pd

def update_inventory_purchase(item_code, quantity):
    """
    매입 관리에서 데이터가 추가 혹은 수정될 때 호출되어 인벤토리를 업데이트합니다.
    :param item_code: 상품 코드
    :param quantity: 매입된 수량
    """
    try:
        inventory_df = pd.read_excel('inventory.xlsx', sheet_name='inventory')
    except FileNotFoundError:
        print("파일을 찾을 수 없습니다: inventory.xlsx")
        return

    # 인벤토리에서 해당 상품의 재고 수량 업데이트
    if item_code in inventory_df['상품 코드'].values:
        idx = inventory_df.index[inventory_df['상품 코드'] == item_code].tolist()[0]
        current_quantity = inventory_df.loc[idx, '수량']
        new_quantity = current_quantity + quantity
        inventory_df.loc[idx, '수량'] = new_quantity
    else:
        # 인벤토리에 해당 상품 코드가 없으면 새로 추가
        new_row = {'상품 코드': item_code, '수량': quantity}
        inventory_df = inventory_df.append(new_row, ignore_index=True)

    # 업데이트된 인벤토리 저장
    inventory_df.to_excel('inventory.xlsx', index=False, sheet_name='inventory')

    # ps.xlsx 파일의 재고 부분 업데이트
    update_ps_stock(inventory_df)

def update_ps_stock(inventory_df):
    """
    ps.xlsx 파일의 재고 부분을 업데이트합니다.
    :param inventory_df: 인벤토리 데이터 프레임
    """
    try:
        ps_df = pd.read_excel('ps.xlsx', sheet_name='ps')
    except FileNotFoundError:
        print("파일을 찾을 수 없습니다: ps.xlsx")
        return

    # 월별 판매가격 * 수량 계산하여 ps.xlsx의 재고 부분에 추가
    for month in range(1, 13):
        month_name = f"{month}월"
        total_sales_price = 0
        for index, row in inventory_df.iterrows():
            item_code = row['상품 코드']
            quantity = row['수량']
            item_price = get_item_price(item_code)  # 상품의 판매 가격 가져오기
            total_sales_price += item_price * quantity

        # 이미 해당 월이 존재하면 업데이트, 없으면 추가
        if month_name in ps_df.columns:
            ps_df.loc[1, month_name] = total_sales_price
        else:
            ps_df[month_name] = total_sales_price

    # 업데이트된 ps.xlsx 파일 저장
    ps_df.to_excel('ps.xlsx', index=False, sheet_name='ps')

def get_item_price(item_code):
    """
    인벤토리에서 특정 상품의 판매 가격을 가져옵니다.
    :param item_code: 상품 코드
    :return: 상품의 판매 가격 (가격이 없으면 0을 반환)
    """
    try:
        inventory_df = pd.read_excel('inventory.xlsx', sheet_name='inventory')
    except FileNotFoundError:
        print("파일을 찾을 수 없습니다: inventory.xlsx")
        return 0

    if item_code in inventory_df['상품 코드'].values:
        idx = inventory_df.index[inventory_df['상품 코드'] == item_code].tolist()[0]
        item_price = inventory_df.loc[idx, '판매 가격']
        return item_price
    else:
        print(f"상품 코드 {item_code}에 해당하는 판매 가격이 없습니다.")
        return 0

def update_inventory_sales(sales_data):
    """
    매출 관리에서 데이터가 추가 혹은 수정될 때 호출되어 인벤토리를 업데이트합니다.
    :param sales_data: 매출 데이터 (상품 코드와 판매 수량을 포함한 딕셔너리 리스트)
    """
    try:
        inventory_df = pd.read_excel('inventory.xlsx', sheet_name='inventory')
    except FileNotFoundError:
        print("파일을 찾을 수 없습니다: inventory.xlsx")
        return

    for data in sales_data:
        item_code = data['상품 코드']
        sales_quantity = data['판매 수량']

        if item_code in inventory_df['상품 코드'].values:
            idx = inventory_df.index[inventory_df['상품 코드'] == item_code].tolist()[0]
            current_quantity = inventory_df.loc[idx, '수량']
            new_quantity = current_quantity - sales_quantity
            inventory_df.loc[idx, '수량'] = new_quantity
        else:
            print(f"인벤토리에 상품 코드 {item_code}가 없습니다.")

    # 업데이트된 인벤토리 저장
    inventory_df.to_excel('inventory.xlsx', index=False, sheet_name='inventory')

    # ps.xlsx 파일의 재고 부분 업데이트
    update_ps_stock(inventory_df)

if __name__ == '__main__':
    # control.py가 직접 실행될 때 테스트용 코드
    update_inventory_purchase('001', 100)  # 매입 추가 예시
    update_inventory_sales([{'상품 코드': '001', '판매 수량': 50}])  # 매출 추가 예시
