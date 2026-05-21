from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import random

app = FastAPI()

# Mở cửa cho Frontend từ GitHub Pages gọi sang mà không bị chặn
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {"message": "Hệ thống Backend Python đã online và sẵn sàng nhận lệnh!"}

# Endpoint 1: Xử lý tính toán (Frontend đang gọi cái này)
@app.get("/api/tinh-toan")
def process_data(so_a: int = 0, so_b: int = 0):
    ket_qua = so_a + so_b
    loi_nhan = f"Backend báo cáo: Tính toán hoàn tất. {so_a} + {so_b} = {ket_qua}."
    return {"status": "success", "result": ket_qua, "message": loi_nhan}

# Endpoint 2: Mở rộng tầm nhìn - API Tự động hóa khối lượng lớn
@app.get("/api/tao-hop-dong")
def tao_hop_dong(ten_doi_tac: str = "Khách hàng mặc định"):
    # Chỗ này sau này cậu có thể nhét pandas, docxtpl vào để gen file hàng loạt
    ma_hd = f"HD-GAO-{random.randint(1000, 9999)}"
    loi_nhan = f"Đã khởi tạo tiến trình xuất hợp đồng {ma_hd} cho đối tác {ten_doi_tac}. Tự động hóa muôn năm!"
    return {"status": "success", "message": loi_nhan}
