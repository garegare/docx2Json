//! EMF（Enhanced Metafile）→ PNG 変換モジュール
//!
//! Windows ビルドのみ GDI32 FFI（`windows-sys` クレート経由）を使用して変換を実施する。
//! 非 Windows では `emf_to_png` は常に `None` を返す。

/// バイト列が EMF 形式かどうかをマジックバイトで判定する
///
/// EMF ファイルの先頭 4 バイトは EMR_HEADER レコードタイプ（= 1）の
/// リトルエンディアン u32 表現（`01 00 00 00`）。
pub fn is_emf(data: &[u8]) -> bool {
    matches!(data, [0x01, 0x00, 0x00, 0x00, ..])
}

/// EMF バイト列を PNG に変換する
///
/// # Windows ビルド
/// GDI32 の `PlayEnhMetaFile` を使ってメモリ DC に描画し、
/// BGRA ビットマップを RGB に変換後 PNG エンコードして返す。
///
/// キャンバスサイズは EMF ヘッダーの `rclFrame`（0.01mm 単位）から
/// 96 DPI 相当のピクセル数を算出する。`rclFrame` がゼロの場合は
/// `rclBounds`（デバイス単位）にフォールバックする。
///
/// # 非 Windows ビルド
/// 常に `None` を返す。
#[cfg(target_os = "windows")]
pub fn emf_to_png(emf_bytes: &[u8]) -> Option<Vec<u8>> {
    use windows_sys::Win32::Foundation::RECT;
    use windows_sys::Win32::Graphics::Gdi::{
        CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteEnhMetaFile, DeleteObject,
        GetEnhMetaFileHeader, PlayEnhMetaFile, SelectObject, SetEnhMetaFileBits,
        BITMAPINFO, BITMAPINFOHEADER, DIB_RGB_COLORS, ENHMETAHEADER,
    };

    const MAX_DIM: u32 = 4096; // 長辺の上限ピクセル数

    unsafe {
        // 1. メモリ上の EMF バイト列からハンドルを作成
        let hemf = SetEnhMetaFileBits(emf_bytes.len() as u32, emf_bytes.as_ptr());
        if hemf.is_null() {
            return None;
        }

        // 2. EMF ヘッダーを取得してキャンバスサイズを決定
        let mut header: ENHMETAHEADER = std::mem::zeroed();
        let sz = std::mem::size_of::<ENHMETAHEADER>() as u32;
        if GetEnhMetaFileHeader(hemf, sz, &mut header) == 0 {
            DeleteEnhMetaFile(hemf);
            return None;
        }

        // rclFrame（0.01mm 単位）→ 96 DPI ピクセル
        let frame_w = (header.rclFrame.right  - header.rclFrame.left).unsigned_abs();
        let frame_h = (header.rclFrame.bottom - header.rclFrame.top ).unsigned_abs();
        let (mut width, mut height) = if frame_w > 0 && frame_h > 0 {
            let px_w = ((frame_w as f64 / 100.0 / 25.4) * 96.0).round() as u32;
            let px_h = ((frame_h as f64 / 100.0 / 25.4) * 96.0).round() as u32;
            (px_w.max(1), px_h.max(1))
        } else {
            // rclBounds（デバイス単位）にフォールバック
            let bw = (header.rclBounds.right  - header.rclBounds.left).unsigned_abs();
            let bh = (header.rclBounds.bottom - header.rclBounds.top ).unsigned_abs();
            (bw.max(64), bh.max(64))
        };

        // 長辺が MAX_DIM を超える場合はアスペクト比を保ってリサイズ
        if width > MAX_DIM || height > MAX_DIM {
            let scale = MAX_DIM as f64 / width.max(height) as f64;
            width  = ((width  as f64 * scale).round() as u32).max(1);
            height = ((height as f64 * scale).round() as u32).max(1);
        }

        // 3. メモリ DC を作成
        let hdc = CreateCompatibleDC(std::ptr::null_mut());
        if hdc.is_null() {
            DeleteEnhMetaFile(hemf);
            return None;
        }

        // 4. Top-down 32bpp DIB を作成
        let bmi = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize:          std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth:         width as i32,
                biHeight:        -(height as i32), // 負値 = top-down DIB
                biPlanes:        1,
                biBitCount:      32,
                biCompression:   0, // BI_RGB
                biSizeImage:     0,
                biXPelsPerMeter: 0,
                biYPelsPerMeter: 0,
                biClrUsed:       0,
                biClrImportant:  0,
            },
            bmiColors: [std::mem::zeroed()],
        };

        let mut bits_ptr: *mut std::ffi::c_void = std::ptr::null_mut();
        let hbitmap = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &mut bits_ptr, std::ptr::null_mut(), 0);
        if hbitmap.is_null() || bits_ptr.is_null() {
            DeleteDC(hdc);
            DeleteEnhMetaFile(hemf);
            return None;
        }

        // DIB 背景を白（0xFF）で初期化（GDI はゼロ初期化 = 黒）
        let pixel_bytes = (width * height * 4) as usize;
        std::ptr::write_bytes(bits_ptr as *mut u8, 0xFF, pixel_bytes);

        let old_obj = SelectObject(hdc, hbitmap);

        // 5. EMF を描画
        let rect = RECT {
            left:   0,
            top:    0,
            right:  width  as i32,
            bottom: height as i32,
        };
        PlayEnhMetaFile(hdc, hemf, &rect);

        // 6. BGRA → RGB（GDI の 32bpp は B/G/R/reserved の順）
        let raw = std::slice::from_raw_parts(bits_ptr as *const u8, pixel_bytes);
        let mut rgb = Vec::with_capacity((width * height * 3) as usize);
        for chunk in raw.chunks_exact(4) {
            rgb.push(chunk[2]); // R
            rgb.push(chunk[1]); // G
            rgb.push(chunk[0]); // B
        }

        // 7. クリーンアップ
        SelectObject(hdc, old_obj);
        DeleteObject(hbitmap);
        DeleteDC(hdc);
        DeleteEnhMetaFile(hemf);

        // 8. image クレートで PNG エンコード
        let img = image::RgbImage::from_raw(width, height, rgb)?;
        let mut png = Vec::new();
        image::DynamicImage::ImageRgb8(img)
            .write_to(
                &mut std::io::Cursor::new(&mut png),
                image::ImageFormat::Png,
            )
            .ok()?;
        Some(png)
    }
}

#[cfg(not(target_os = "windows"))]
pub fn emf_to_png(_emf_bytes: &[u8]) -> Option<Vec<u8>> {
    None
}
