"""
マイソク（販売図面）Excel生成スクリプト
テンプレートファイルをコピーして値を書き込む方式で完全再現
フォント：HGS明朝B、カラー：#002060、構造：85列×35行

テンプレートのセル構造（調査済み）:
  X1:BG2   - 物件名（フォント18pt, color=FF002060）
  BP1:CC2  - 価格（フォント24pt, color=FF002060, 中央寄せ）
  CD1:CG2  - 「万円」（固定テキスト、変更不要）
  BP3:CG4  - 所在地（フォント12pt, color=FF002060）
  BP5:CG7  - 交通（フォント9pt, color=FF002060）
  BP8:CG8  - 専有面積（面積+㎡で書き込み）
  BP9:CG9  - バルコニー面積
  BP10:CG10 - 室内洗濯機置場
  BP11:CG11 - 土地権利
  BP12:CG12 - 構造（フォント11pt）
  BP13:CG13 - 総戸数（数値+「戸」で書き込み）
  BP14:CG14 - 築年月（文字列で書き込み）
  BP15:CG15 - 現況
  BP16:CG16 - 引渡日
  BP17:CG17 - 管理会社（フォント11pt）
  BP18:BT18 - 管理費（数値、右寄せ）
  BP19:BT19 - 修繕積立金（数値、右寄せ）
  BP20:BT20 - その他費用（数値、右寄せ）
  CA18:CE20 - 合計（数式 =BP18+BP19+BP20 は既存）
  BP21:CG22 - 賃料（フォント18pt）
  BP23:CG28 - その他・備考（フォント10pt, wrap=True）
  A31:AJ31  - 免許番号（背景色=FF002060, 白文字）
  A32:AJ34  - 会社名（背景色=FF002060, 白文字）
  AK31:BF35 - 電話・FAX（フォント18pt, color=FF002060）
  BG31:BR35 - 担当者（フォント18pt, color=FF002060）
  CA31:CG32 - 取引形態（フォント14pt）
  CA33:CG34 - 手数料（フォント14pt）
  A35:AJ35  - 会社住所（背景色=FF002060, 白文字）
"""
import openpyxl
import io
import os

# テンプレートファイルのパス（app.pyと同じディレクトリ）
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'maisoku_template.xlsx')


def _set_value_only(ws, coord, value):
    """セルの値のみを設定（スタイルは一切変更しない）"""
    cell = ws[coord]
    cell.value = value


def generate_maisoku(data: dict) -> bytes:
    """
    テンプレートをコピーして物件情報を書き込み、Excelバイナリを返す

    data keys:
      prop_name       : 物件名
      floor           : 階数表示（例：「7階部分」）
      price           : 販売価格（万円）
      location        : 所在地
      station         : 交通（最寄駅・徒歩分数）
      area            : 専有面積（㎡）
      balcony         : バルコニー面積
      washer          : 室内洗濯機置場（有/無）
      land_rights     : 土地権利
      structure       : 構造（例：鉄筋コンクリート造）
      floor_num       : 総階数
      room_floor      : 部屋の階
      total_units     : 総戸数
      built_date      : 築年月（例：2020年10月）
      status          : 現況
      delivery        : 引渡日
      mgmt_co         : 管理会社
      mgmt_fee        : 管理費（円）
      repair_fund     : 修繕積立金（円）
      other_fee       : その他費用（円）
      rental          : 賃料
      notes           : その他・備考
      company_name    : 会社名
      license_no      : 免許番号
      address_company : 会社住所
      tel             : 電話番号
      fax             : FAX番号
      staff           : 担当者名
      trade_type      : 取引形態
      commission      : 手数料
    """
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {TEMPLATE_PATH}")

    # テンプレートをロード（スタイル・数式を保持）
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # ===== 物件名（X1:BG2 マージ済み）=====
    prop_name = data.get('prop_name', '')
    floor_str = data.get('floor', '')
    title_text = f"{prop_name}　{floor_str}" if floor_str else prop_name
    _set_value_only(ws, 'X1', title_text)

    # ===== 価格（BP1:CC2 マージ済み）=====
    price = data.get('price', '')
    _set_value_only(ws, 'BP1', price)

    # ===== 所在地（BP3:CG4 マージ済み）=====
    _set_value_only(ws, 'BP3', data.get('location', ''))

    # ===== 交通（BP5:CG7 マージ済み）=====
    _set_value_only(ws, 'BP5', data.get('station', ''))

    # ===== 専有面積（BP8:CG8 マージ済み）=====
    # テンプレートには「㎡」という値が入っているが、面積の数値+㎡に置き換える
    area = data.get('area', '')
    area_text = f"{area}㎡" if area and '㎡' not in str(area) else str(area)
    _set_value_only(ws, 'BP8', area_text)

    # ===== バルコニー面積（BP9:CG9 マージ済み）=====
    _set_value_only(ws, 'BP9', data.get('balcony', '-'))

    # ===== 室内洗濯機置場（BP10:CG10 マージ済み）=====
    _set_value_only(ws, 'BP10', data.get('washer', '有'))

    # ===== 土地権利（BP11:CG11 マージ済み）=====
    _set_value_only(ws, 'BP11', data.get('land_rights', '所有権'))

    # ===== 構造（BP12:CG12 マージ済み）=====
    floor_num = data.get('floor_num', '')
    room_floor = data.get('room_floor', '')
    if floor_num and room_floor:
        structure_text = f"鉄筋コンクリート造地上　{floor_num}階建　{room_floor}階部分"
    elif floor_num:
        structure_text = f"鉄筋コンクリート造地上　{floor_num}階建"
    else:
        structure_text = data.get('structure', '鉄筋コンクリート造地上　　階建　　階部分')
    _set_value_only(ws, 'BP12', structure_text)

    # ===== 総戸数（BP13:CG13 マージ済み）=====
    # テンプレートには「戸」が入っているので、数値+「戸」に置き換える
    total_units = data.get('total_units', '')
    if total_units:
        units_text = f"{total_units}戸" if '戸' not in str(total_units) else str(total_units)
    else:
        units_text = '戸'
    _set_value_only(ws, 'BP13', units_text)

    # ===== 築年月（BP14:CG14 マージ済み）=====
    # テンプレートには日付シリアル値が入っているので文字列で上書き
    built_date = data.get('built_date', '')
    _set_value_only(ws, 'BP14', built_date)

    # ===== 現況（BP15:CG15 マージ済み）=====
    _set_value_only(ws, 'BP15', data.get('status', '賃貸中'))

    # ===== 引渡日（BP16:CG16 マージ済み）=====
    _set_value_only(ws, 'BP16', data.get('delivery', '要相談'))

    # ===== 管理会社（BP17:CG17 マージ済み）=====
    _set_value_only(ws, 'BP17', data.get('mgmt_co', ''))

    # ===== 管理費（BP18:BT18 マージ済み）=====
    mgmt_fee = data.get('mgmt_fee', '')
    try:
        mgmt_fee_val = int(str(mgmt_fee).replace(',', '').replace('円', '').strip())
    except:
        mgmt_fee_val = mgmt_fee if mgmt_fee else None
    _set_value_only(ws, 'BP18', mgmt_fee_val)

    # ===== 修繕積立金（BP19:BT19 マージ済み）=====
    repair_fund = data.get('repair_fund', '')
    try:
        repair_fund_val = int(str(repair_fund).replace(',', '').replace('円', '').strip())
    except:
        repair_fund_val = repair_fund if repair_fund else None
    _set_value_only(ws, 'BP19', repair_fund_val)

    # ===== その他費用（BP20:BT20 マージ済み）=====
    other_fee = data.get('other_fee', '')
    try:
        other_fee_val = int(str(other_fee).replace(',', '').replace('円', '').strip())
    except:
        other_fee_val = other_fee if other_fee else None
    _set_value_only(ws, 'BP20', other_fee_val)

    # ===== 合計（CA18:CE20）は既存の数式 =BP18+BP19+BP20 をそのまま使用 =====

    # ===== 賃料（BP21:CG22 マージ済み）=====
    rental = data.get('rental', '')
    _set_value_only(ws, 'BP21', rental)

    # ===== その他・備考（BP23:CG28 マージ済み）=====
    notes = data.get('notes', '※管理費・修繕積立金の内訳は調査中')
    _set_value_only(ws, 'BP23', notes)

    # ===== 免許番号（A31:AJ31 マージ済み）=====
    license_no = data.get('license_no', '国土交通大臣（１）第９９２１号')
    _set_value_only(ws, 'A31', license_no)

    # ===== 会社名（A32:AJ34 マージ済み）=====
    company_name = data.get('company_name', '株式会社ＧＲエステート')
    _set_value_only(ws, 'A32', company_name)

    # ===== 電話・FAX（AK31:BF35 マージ済み）=====
    tel = data.get('tel', '03-6432-5088')
    fax = data.get('fax', '03-6432-5087')
    _set_value_only(ws, 'AK31', f"電話番号：{tel}\nFAX番号：{fax}")

    # ===== 担当者（BG31:BR35 マージ済み）=====
    staff = data.get('staff', '野村')
    _set_value_only(ws, 'BG31', f"担当：{staff}")

    # ===== 取引形態（CA31:CG32 マージ済み）=====
    _set_value_only(ws, 'CA31', data.get('trade_type', '媒介'))

    # ===== 手数料（CA33:CG34 マージ済み）=====
    _set_value_only(ws, 'CA33', data.get('commission', '分かれ'))

    # ===== 会社住所（A35:AJ35 マージ済み）=====
    address_company = data.get('address_company', '〒141-0022　東京都品川区東五反田1-16-11-4階')
    _set_value_only(ws, 'A35', address_company)

    # ===== 印刷設定 =====
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = 'A1:CG35'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
