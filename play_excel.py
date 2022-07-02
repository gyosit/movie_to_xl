import xlwings as xw
import cv2

if __name__ == "__main__":
    LENGTH = 10 # 行の高さ・列の幅
    HEIGHT_COL_RATIO = 10 # 列幅は行高の1/10くらいだと正方形に見える
    SIZE_RATIO = 10 # 画像サイズの縮小率
    FRAME_RATE = 10 # 描画するフレームの割合 (大きいほどカクつくが再生速度は遅くなる)

    EX_TITLE = "book1" # 書き込み先のエクセルのタイトル (あらかじめ開いておく)
    SH_TITLE = "sheet1" # 書き込み先のシート名
    MOVIE_TITLE = "./つむぎbb.mp4" # 変換元の動画タイトル

    wb = xw.Book(EX_TITLE) # 新しく開いたブック
    sht = wb.sheets[SH_TITLE] # 操作するシート

    rectange_cells = None
    h, w = None, None

    cap = cv2.VideoCapture(MOVIE_TITLE)

    i = 0
    
    while True:
        ret, img = cap.read()

        if ret == False:
            break

        if i == 0:
            h, w = img.shape[:2]
            h, w = int(h / SIZE_RATIO), int(w / SIZE_RATIO)
            rectange_cells = sht.range((1, 1), (h, w))
            rectange_cells.row_height = LENGTH
            rectange_cells.column_width = LENGTH  / HEIGHT_COL_RATIO
        
        i += 1

        if i % FRAME_RATE != 0:
            continue

        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        dst = 255 - cv2.resize(img, dsize=(w, h))
        
        dst[0] = 0
        dst[-1] = 255
        rectange_cells.value = dst
