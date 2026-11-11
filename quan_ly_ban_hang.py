
import os
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import datetime, timedelta

import openpyxl
import xlsxwriter

THU_MUC = os.path.dirname(os.path.abspath(__file__))
FILE_EXCEL = os.path.join(THU_MUC, "danh_sach_1000_san_pham.xlsx")
THU_MUC_XUAT = os.path.join(THU_MUC, "xuat")
MONGO_URI = os.environ.get("MONGO_URI", "mongodb://localhost:27017")
MONGO_DB = os.environ.get("MONGO_DB", "quan_ly_ban_hang")
CO_SO_TRANG = 20

from pymongo import MongoClient, ReturnDocument

AFTER = ReturnDocument.AFTER


def ket_noi():
    client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=3000)
    try:
        client.admin.command("ping")
    except Exception as e:
        sys.exit(f"Không kết nối được MongoDB tại {MONGO_URI}: {e}")
    db = client[MONGO_DB]
    db.products.create_index("ma_sp", unique=True)
    db.orders.create_index("ma_don", unique=True)
    db.orders.create_index("ngay")
    return db


def bo_dau(s):
    s = unicodedata.normalize("NFD", str(s).replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def tien(x):
    return f"{x:,.0f}".replace(",", ".")


def doc_tien(s):
    s = re.sub(r"(?i)vn[đd]|đ|[\s.,]", "", s)
    return int(s) if s.isdigit() else None


def hoi(prompt, mac_dinh=None):
    s = input(prompt + (f" [{mac_dinh}]" if mac_dinh else "") + ": ").strip()
    return s or (mac_dinh or "")


def hoi_so(prompt, toi_thieu=1):
    while True:
        s = hoi(prompt)
        if s == "":
            return None
        v = doc_tien(s)
        if v is not None and v >= toi_thieu:
            return v
        print(f"  Giá trị không hợp lệ (cần số nguyên >= {toi_thieu}).")


def hoi_ngay(prompt):
    while True:
        s = hoi(prompt + " (dd/mm/yyyy, Enter = bỏ qua)")
        if not s:
            return None
        try:
            return datetime.strptime(s, "%d/%m/%Y")
        except ValueError:
            print("  Ngày không hợp lệ.")


def bang(dong, tieu_de, can_phai=()):
    """In bảng văn bản. can_phai: chỉ số các cột căn phải."""
    dong = [[str(c) for c in r] for r in dong]
    rong = [max(len(h), *(len(r[i]) for r in dong)) if dong else len(h)
            for i, h in enumerate(tieu_de)]
    def fmt(r):
        return " | ".join(c.rjust(rong[i]) if i in can_phai else c.ljust(rong[i])
                          for i, c in enumerate(r))
    print(fmt(tieu_de))
    print("-+-".join("-" * w for w in rong))
    for r in dong:
        print(fmt(r))


def phan_trang(tong, ve_trang):
    """Điều hướng trang. ve_trang(skip, limit) in một trang."""
    trang, so_trang = 0, max(1, -(-tong // CO_SO_TRANG))
    while True:
        ve_trang(trang * CO_SO_TRANG, CO_SO_TRANG)
        print(f"\nTrang {trang + 1}/{so_trang} - tổng {tong} bản ghi. "
              "[n] sau  [p] trước  [số] tới trang  [q] thoát")
        c = input("> ").strip().lower()
        if c == "n" and trang + 1 < so_trang: trang += 1
        elif c == "p" and trang > 0: trang -= 1
        elif c.isdigit() and 1 <= int(c) <= so_trang: trang = int(c) - 1
        elif c in ("q", ""): return


def doc_excel(duong_dan=FILE_EXCEL):
    """Chức năng 1: đọc sản phẩm từ Excel."""
    if not os.path.exists(duong_dan):
        raise FileNotFoundError(duong_dan)
    ws = openpyxl.load_workbook(duong_dan, read_only=True)["Danh sách sản phẩm"]
    ds = []
    for r in ws.iter_rows(min_row=2, max_col=10, values_only=True):
        if not r[0]:
            continue
        nfc = lambda v: unicodedata.normalize("NFC", str(v).strip())
        sp = {
            "ma_sp": nfc(r[0]), "ten_sp": nfc(r[1]), "danh_muc": nfc(r[2]),
            "thuong_hieu": nfc(r[3]), "don_vi": nfc(r[4]),
            "don_gia": int(r[5]), "ton_kho": int(r[6]),
            "ban_thang": int(r[7]), "ban_quy": int(r[8]), "ban_nam": int(r[9]),
        }
        sp["tim_kiem"] = bo_dau(f"{sp['ma_sp']} {sp['ten_sp']} {sp['danh_muc']} {sp['thuong_hieu']}")
        ds.append(sp)
    return ds


def nap_mongodb(db):
    """Chức năng 2: nạp sản phẩm vào MongoDB."""
    path = hoi("Đường dẫn file Excel", FILE_EXCEL)
    try:
        ds = doc_excel(path)
    except Exception as e:
        print("Không đọc được file:", e); return
    print(f"Đã đọc {len(ds)} sản phẩm từ Excel.")
    if db.products.count_documents({}):
        ghi_de = hoi("CSDL đã có sản phẩm. Xóa hết và nạp lại? (y/N)", "n").lower() == "y"
        if ghi_de:
            db.products.delete_many({})
        else:
            co = {d["ma_sp"] for d in db.products.find({})}
            ds = [d for d in ds if d["ma_sp"] not in co]
            print(f"Bỏ qua sản phẩm đã tồn tại, chỉ thêm mới {len(ds)} sản phẩm.")
    if ds:
        db.products.insert_many(ds)
    print(f"Hoàn tất. Hiện có {db.products.count_documents({})} sản phẩm trong CSDL.")


def in_sp(ds):
    bang([[s["ma_sp"], s["ten_sp"], s["danh_muc"], s["thuong_hieu"], s["don_vi"],
           tien(s["don_gia"]), s["ton_kho"]] for s in ds],
         ["Mã", "Tên sản phẩm", "Danh mục", "Thương hiệu", "ĐVT", "Đơn giá (VNĐ)", "Tồn"],
         can_phai=(5, 6))


def hien_thi_sp(db, bo_loc=None):
    """Chức năng 3: hiển thị danh sách sản phẩm (có phân trang)."""
    bo_loc = bo_loc or {}
    tong = db.products.count_documents(bo_loc)
    if not tong:
        print("Không có sản phẩm nào."); return
    phan_trang(tong, lambda skip, lim: in_sp(
        list(db.products.find(bo_loc).sort("ma_sp", 1).skip(skip).limit(lim))))


def tim_kiem_sp(db):
    """Chức năng 4: tìm theo từ khóa (không phân biệt hoa/thường, dấu), danh mục, khoảng giá."""
    tu = hoi("Từ khóa (mã/tên/danh mục/thương hiệu, Enter = bỏ qua)")
    dm = hoi("Danh mục (Enter = tất cả)")
    gia_tu = hoi_so("Giá từ (VNĐ, Enter = bỏ qua)", 0)
    gia_den = hoi_so("Giá đến (VNĐ, Enter = bỏ qua)", 0)
    bo_loc = {}
    tach = [re.escape(w) for w in bo_dau(tu).split()]
    if tach:
        bo_loc["tim_kiem"] = {"$regex": "^" + "".join(f"(?=.*{w})" for w in tach)}
    if dm:
        bo_loc["danh_muc"] = {"$regex": "^" + re.escape(dm) + "$", "$options": "i"}
    gia = {}
    if gia_tu is not None: gia["$gte"] = gia_tu
    if gia_den is not None: gia["$lte"] = gia_den
    if gia: bo_loc["don_gia"] = gia
    hien_thi_sp(db, bo_loc)


def tim_sp(db, ma):
    return db.products.find_one({"ma_sp": ma.strip().upper()})


def ma_sp_tiep_theo(db):
    so = [int(m.group(1)) for d in db.products.find({})
          if (m := re.fullmatch(r"SP(\d+)", d["ma_sp"]))]
    return f"SP{(max(so) + 1 if so else 1):04d}"


def them_sp(db):
    """Chức năng 5: thêm sản phẩm mới."""
    ma = hoi("Mã sản phẩm (Enter = tự sinh)").upper() or ma_sp_tiep_theo(db)
    if tim_sp(db, ma):
        print(f"Mã {ma} đã tồn tại."); return
    ten = hoi("Tên sản phẩm")
    if not ten:
        print("Tên không được để trống."); return
    dm, th = hoi("Danh mục"), hoi("Thương hiệu")
    dv = hoi("Đơn vị tính", "Cái")
    gia = hoi_so("Đơn giá (VNĐ)", 1)
    ton = hoi_so("Tồn kho ban đầu", 0) or 0
    if gia is None:
        print("Cần nhập đơn giá."); return
    db.products.insert_one({
        "ma_sp": ma, "ten_sp": ten, "danh_muc": dm, "thuong_hieu": th, "don_vi": dv,
        "don_gia": gia, "ton_kho": ton, "ban_thang": 0, "ban_quy": 0, "ban_nam": 0,
        "tim_kiem": bo_dau(f"{ma} {ten} {dm} {th}"),
    })
    print(f"Đã thêm {ma} - {ten} ({tien(gia)} VNĐ).")


def cap_nhat_gia(db):
    """Chức năng 6: cập nhật giá sản phẩm."""
    sp = tim_sp(db, hoi("Mã sản phẩm"))
    if not sp:
        print("Không tìm thấy sản phẩm."); return
    print(f"{sp['ma_sp']} - {sp['ten_sp']}: giá hiện tại {tien(sp['don_gia'])} VNĐ")
    gia = hoi_so("Giá mới (VNĐ)", 1)
    if gia is None:
        print("Hủy."); return
    db.products.update_one({"ma_sp": sp["ma_sp"]},
                           {"$set": {"don_gia": gia, "gia_cap_nhat": datetime.now()}})
    print(f"Đã đổi giá: {tien(sp['don_gia'])} -> {tien(gia)} VNĐ. "
          "(Đơn đã lập trước đó giữ nguyên giá cũ.)")


def sinh_ma_don(db):
    hom_nay = datetime.now().strftime("%Y%m%d")
    c = db.counters.find_one_and_update(
        {"_id": "don_" + hom_nay}, {"$inc": {"seq": 1}}, upsert=True, return_document=AFTER)
    return f"DH{hom_nay}-{c['seq']:04d}"


def lap_don_hang(db):
    """Chức năng 7: lập đơn hàng (kiểm tra và trừ tồn kho)."""
    khach = hoi("Tên khách hàng", "Khách lẻ")
    sdt = hoi("Số điện thoại (Enter = bỏ qua)")
    gio = {} 
    print("Nhập từng sản phẩm. Để trống mã sản phẩm để kết thúc.")
    while True:
        ma = hoi("Mã sản phẩm").upper()
        if not ma:
            break
        sp = tim_sp(db, ma)
        if not sp:
            print("  Không tìm thấy sản phẩm."); continue
        da_co = gio[ma]["so_luong"] if ma in gio else 0
        con = sp["ton_kho"] - da_co
        print(f"  {sp['ten_sp']} - {tien(sp['don_gia'])} VNĐ/{sp['don_vi']} (còn {con})")
        if con <= 0:
            print("  Hết hàng."); continue
        sl = hoi_so("  Số lượng", 1)
        if sl is None:
            continue
        if sl > con:
            print(f"  Không đủ hàng (chỉ còn {con})."); continue
        if ma in gio:
            gio[ma]["so_luong"] += sl
            gio[ma]["thanh_tien"] = gio[ma]["so_luong"] * gio[ma]["don_gia"]
        else:
            gio[ma] = {"ma_sp": ma, "ten_sp": sp["ten_sp"], "danh_muc": sp["danh_muc"],
                       "don_vi": sp["don_vi"], "don_gia": sp["don_gia"],
                       "so_luong": sl, "thanh_tien": sl * sp["don_gia"]}
        print(f"  Tạm tính: {tien(sum(i['thanh_tien'] for i in gio.values()))} VNĐ")
    if not gio:
        print("Đơn hàng trống, đã hủy."); return
    items = list(gio.values())
    tong = sum(i["thanh_tien"] for i in items)
    print(f"\nTổng {len(items)} mặt hàng, thành tiền {tien(tong)} VNĐ.")
    if hoi("Xác nhận lập đơn? (Y/n)", "y").lower() != "y":
        print("Đã hủy đơn."); return

    da_tru = []
    for it in items:
        r = db.products.update_one(
            {"ma_sp": it["ma_sp"], "ton_kho": {"$gte": it["so_luong"]}},
            {"$inc": {"ton_kho": -it["so_luong"], "ban_thang": it["so_luong"],
                      "ban_quy": it["so_luong"], "ban_nam": it["so_luong"]}})
        if r.modified_count != 1:
            for d in da_tru:
                db.products.update_one({"ma_sp": d["ma_sp"]}, {"$inc": {
                    "ton_kho": d["so_luong"], "ban_thang": -d["so_luong"],
                    "ban_quy": -d["so_luong"], "ban_nam": -d["so_luong"]}})
            print(f"Lỗi: {it['ma_sp']} vừa hết hàng. Đơn chưa được lập."); return
        da_tru.append(it)
    don = {"ma_don": sinh_ma_don(db), "ngay": datetime.now(), "khach_hang": khach,
           "sdt": sdt, "items": items, "tong_tien": tong}
    db.orders.insert_one(don)
    print(f"Đã lập đơn {don['ma_don']}.\n")
    in_hoa_don(don)


def in_hoa_don(d):
    r = 78
    print("=" * r)
    print("HÓA ĐƠN BÁN HÀNG".center(r))
    print("=" * r)
    print(f"Số hóa đơn : {d['ma_don']}")
    print(f"Ngày lập   : {d['ngay']:%d/%m/%Y %H:%M}")
    print(f"Khách hàng : {d['khach_hang']}" + (f" - {d['sdt']}" if d.get("sdt") else ""))
    print("-" * r)
    bang([[i + 1, x["ma_sp"], x["ten_sp"], x["don_vi"], tien(x["don_gia"]),
           x["so_luong"], tien(x["thanh_tien"])] for i, x in enumerate(d["items"])],
         ["STT", "Mã", "Tên sản phẩm", "ĐVT", "Đơn giá", "SL", "Thành tiền"], can_phai=(0, 4, 5, 6))
    print("-" * r)
    print(f"TỔNG CỘNG ({sum(x['so_luong'] for x in d['items'])} sản phẩm): "
          f"{tien(d['tong_tien'])} VNĐ".rjust(r))
    print("=" * r)
    print("Cảm ơn quý khách!".center(r))


def hien_thi_hoa_don(db):
    """Chức năng 8: hiển thị hóa đơn theo mã đơn."""
    ma = hoi("Mã đơn hàng (VD: DH20261110-0001)").upper()
    d = db.orders.find_one({"ma_don": ma})
    if d: print(); in_hoa_don(d)
    else: print("Không tìm thấy đơn hàng.")


def bo_loc_ngay(tu, den):
    if not (tu or den): return {}
    n = {}
    if tu: n["$gte"] = tu
    if den: n["$lt"] = den + timedelta(days=1)
    return {"ngay": n}


def xem_don_hang(db):
    """Chức năng 9: xem danh sách đơn hàng (mới nhất trước), lọc theo ngày."""
    bl = bo_loc_ngay(hoi_ngay("Từ ngày"), hoi_ngay("Đến ngày"))
    tong = db.orders.count_documents(bl)
    if not tong:
        print("Chưa có đơn hàng nào."); return
    def ve(skip, lim):
        bang([[d["ma_don"], f"{d['ngay']:%d/%m/%Y %H:%M}", d["khach_hang"],
               sum(i["so_luong"] for i in d["items"]), tien(d["tong_tien"])]
              for d in db.orders.find(bl).sort("ngay", -1).skip(skip).limit(lim)],
             ["Mã đơn", "Ngày", "Khách hàng", "SL", "Tổng tiền (VNĐ)"], can_phai=(3, 4))
    phan_trang(tong, ve)


def thong_ke(db):
    """Chức năng 10: thống kê doanh thu."""
    bl = bo_loc_ngay(hoi_ngay("Từ ngày"), hoi_ngay("Đến ngày"))
    don = list(db.orders.find(bl))
    if not don:
        print("Không có đơn hàng trong khoảng này."); return
    tong = sum(d["tong_tien"] for d in don)
    sl = sum(i["so_luong"] for d in don for i in d["items"])
    print("\n=== TỔNG QUAN ===")
    bang([["Số đơn hàng", len(don)], ["Tổng số lượng bán", sl],
          ["Tổng doanh thu (VNĐ)", tien(tong)],
          ["Doanh thu trung bình/đơn (VNĐ)", tien(tong / len(don))]],
         ["Chỉ tiêu", "Giá trị"], can_phai=(1,))
    def nhom(khoa):
        g = defaultdict(lambda: [0, 0])
        for d in don:
            g[khoa(d)][0] += 1; g[khoa(d)][1] += d["tong_tien"]
        return g
    print("\n=== THEO THÁNG ===")
    g = nhom(lambda d: d["ngay"].strftime("%Y-%m"))
    bang([[k, v[0], tien(v[1])] for k, v in sorted(g.items())],
         ["Tháng", "Số đơn", "Doanh thu (VNĐ)"], can_phai=(1, 2))
    print("\n=== THEO NGÀY (tối đa 31 ngày gần nhất) ===")
    g = nhom(lambda d: d["ngay"].strftime("%Y-%m-%d"))
    bang([[datetime.strptime(k, "%Y-%m-%d").strftime("%d/%m/%Y"), v[0], tien(v[1])]
          for k, v in sorted(g.items())[-31:]],
         ["Ngày", "Số đơn", "Doanh thu (VNĐ)"], can_phai=(1, 2))
    sp = defaultdict(lambda: [0, 0, ""])
    dmuc = defaultdict(lambda: [0, 0])
    for d in don:
        for i in d["items"]:
            s = sp[i["ma_sp"]]; s[0] += i["so_luong"]; s[1] += i["thanh_tien"]; s[2] = i["ten_sp"]
            c = dmuc[i["danh_muc"]]; c[0] += i["so_luong"]; c[1] += i["thanh_tien"]
    print("\n=== TOP 10 SẢN PHẨM THEO DOANH THU ===")
    top = sorted(sp.items(), key=lambda kv: -kv[1][1])[:10]
    bang([[k, v[2], v[0], tien(v[1])] for k, v in top],
         ["Mã", "Tên sản phẩm", "SL bán", "Doanh thu (VNĐ)"], can_phai=(2, 3))
    print("\n=== THEO DANH MỤC ===")
    bang([[k, v[0], tien(v[1]), f"{v[1] / tong:.1%}"]
          for k, v in sorted(dmuc.items(), key=lambda kv: -kv[1][1])],
         ["Danh mục", "SL bán", "Doanh thu (VNĐ)", "Tỷ trọng"], can_phai=(1, 2, 3))
    ls = sum(s["don_gia"] * s["ban_nam"] for s in db.products.find({}))
    print(f"\n(Tham khảo) Doanh thu năm ước tính từ dữ liệu Excel ban đầu: {tien(ls)} VNĐ")


def tao_file(ten):
    os.makedirs(THU_MUC_XUAT, exist_ok=True)
    return os.path.join(THU_MUC_XUAT, f"{ten}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")


def xuat_san_pham(db):
    """Chức năng 11: xuất sản phẩm ra Excel."""
    ds = list(db.products.find({}).sort("ma_sp", 1))
    if not ds:
        print("Chưa có sản phẩm để xuất."); return
    path = tao_file("san_pham")
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Danh sách sản phẩm")
    head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                          "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    so = wb.add_format({"num_format": "#,##0"})
    tieu_de = ["Mã sản phẩm", "Tên sản phẩm", "Danh mục", "Thương hiệu", "Đơn vị tính",
               "Đơn giá (VNĐ)", "Tồn kho", "Số lượng bán tháng", "Số lượng bán quý",
               "Số lượng bán năm", "Tiền thu về năm (VNĐ)"]
    ws.set_row(0, 32)
    for c, h in enumerate(tieu_de):
        ws.write(0, c, h, head)
    for r, s in enumerate(ds, start=1):
        ws.write_row(r, 0, [s["ma_sp"], s["ten_sp"], s["danh_muc"], s["thuong_hieu"], s["don_vi"]])
        for c, k in zip(range(5, 10), ["don_gia", "ton_kho", "ban_thang", "ban_quy", "ban_nam"]):
            ws.write_number(r, c, s[k], so)
        ws.write_formula(r, 10, f"=F{r + 1}*J{r + 1}", so, s["don_gia"] * s["ban_nam"])
    for c, w in enumerate([13, 38, 18, 16, 11, 15, 10, 14, 14, 14, 20]):
        ws.set_column(c, c, w)
    ws.freeze_panes(1, 2)
    ws.autofilter(0, 0, len(ds), 10)
    wb.close()
    print(f"Đã xuất {len(ds)} sản phẩm -> {path}")


def xuat_don_hang(db):
    """Chức năng 12: xuất đơn hàng ra Excel (sheet tổng hợp + sheet chi tiết)."""
    don = list(db.orders.find({}).sort("ngay", 1))
    if not don:
        print("Chưa có đơn hàng để xuất."); return
    path = tao_file("don_hang")
    wb = xlsxwriter.Workbook(path)
    head = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                          "border": 1, "align": "center", "valign": "vcenter"})
    so = wb.add_format({"num_format": "#,##0"})
    ngay = wb.add_format({"num_format": "dd/mm/yyyy hh:mm"})
    dam = wb.add_format({"bold": True, "num_format": "#,##0", "top": 1})
    ws = wb.add_worksheet("Đơn hàng")
    ws.write_row(0, 0, ["Mã đơn", "Ngày", "Khách hàng", "SĐT", "Số dòng", "Tổng SL",
                        "Tổng tiền (VNĐ)"], head)
    for r, d in enumerate(don, start=1):
        ws.write(r, 0, d["ma_don"]); ws.write_datetime(r, 1, d["ngay"], ngay)
        ws.write(r, 2, d["khach_hang"]); ws.write(r, 3, d.get("sdt", ""))
        ws.write_number(r, 4, len(d["items"]), so)
        ws.write_number(r, 5, sum(i["so_luong"] for i in d["items"]), so)
        ws.write_number(r, 6, d["tong_tien"], so)
    n = len(don)
    ws.write(n + 1, 0, "TỔNG", dam)
    ws.write_formula(n + 1, 6, f"=SUM(G2:G{n + 1})", dam, sum(d["tong_tien"] for d in don))
    for c, w in enumerate([22, 18, 24, 14, 10, 10, 18]):
        ws.set_column(c, c, w)
    ws.freeze_panes(1, 0)
    ct = wb.add_worksheet("Chi tiết")
    ct.write_row(0, 0, ["Mã đơn", "Ngày", "Mã SP", "Tên sản phẩm", "Danh mục", "ĐVT",
                        "Đơn giá (VNĐ)", "Số lượng", "Thành tiền (VNĐ)"], head)
    r = 0
    for d in don:
        for i in d["items"]:
            r += 1
            ct.write(r, 0, d["ma_don"]); ct.write_datetime(r, 1, d["ngay"], ngay)
            ct.write_row(r, 2, [i["ma_sp"], i["ten_sp"], i["danh_muc"], i["don_vi"]])
            ct.write_number(r, 6, i["don_gia"], so); ct.write_number(r, 7, i["so_luong"], so)
            ct.write_formula(r, 8, f"=G{r + 1}*H{r + 1}", so, i["thanh_tien"])
    for c, w in enumerate([22, 18, 10, 38, 18, 8, 15, 10, 18]):
        ct.set_column(c, c, w)
    ct.freeze_panes(1, 0)
    ct.autofilter(0, 0, r, 8)
    wb.close()
    print(f"Đã xuất {n} đơn hàng ({r} dòng chi tiết) -> {path}")


MENU = [
    ("Đọc sản phẩm từ Excel (xem thử)", lambda db: print(
        f"Đọc được {len(doc_excel())} sản phẩm từ {os.path.basename(FILE_EXCEL)}")),
    ("Nạp sản phẩm vào MongoDB", nap_mongodb),
    ("Hiển thị danh sách sản phẩm", hien_thi_sp),
    ("Tìm kiếm sản phẩm", tim_kiem_sp),
    ("Thêm sản phẩm mới", them_sp),
    ("Cập nhật giá sản phẩm", cap_nhat_gia),
    ("Lập đơn hàng", lap_don_hang),
    ("Hiển thị hóa đơn", hien_thi_hoa_don),
    ("Xem danh sách đơn hàng", xem_don_hang),
    ("Thống kê doanh thu", thong_ke),
    ("Xuất sản phẩm ra Excel", xuat_san_pham),
    ("Xuất đơn hàng ra Excel", xuat_don_hang),
]


def main():
    db = ket_noi()
    while True:
        print("\n" + "=" * 46)
        print("  QUẢN LÝ ĐƠN HÀNG VÀ BÁN HÀNG")
        print("=" * 46)
        for i, (ten, _) in enumerate(MENU, 1):
            print(f" {i:>2}. {ten}")
        print("  0. Thoát")
        try:
            c = input("Chọn chức năng: ").strip()
            if c == "0":
                break
            if c.isdigit() and 1 <= int(c) <= len(MENU):
                print()
                MENU[int(c) - 1][1](db)
            else:
                print("Lựa chọn không hợp lệ.")
        except (KeyboardInterrupt, EOFError):
            print(); break
    print("Tạm biệt!")


if __name__ == "__main__":
    main()