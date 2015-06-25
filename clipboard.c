/*
* Copyright (C) 1996-2003 by Nakashima Tomoaki. All rights reserved.
*		http://www.nakka.com/
*		nakka@nakka.com
*/

#include "stdafx.h"
#include "clipboard.h"

/*
* bitmap_to_dib - ビットマップをDIBに変換
*/
static BYTE *bitmap_to_dib(const HBITMAP hbmp, size_t *size)
{
	HDC hdc;
	BITMAP bmp;
	BYTE *ret;		// BITMAPINFO
	DWORD biComp = BI_RGB;
	size_t len;
	size_t hsize;
	DWORD err;
	int color_bit = 0;

	// BITMAP情報取得--Information acquisition
	if(GetObject(hbmp, sizeof(BITMAP), &bmp) == 0) {
		return NULL;
	}
	switch(bmp.bmBitsPixel) {
	case 1:
		color_bit = 2;
		break;
	case 4:
		color_bit = 16;
		break;
	case 8:
		color_bit = 256;
		break;
	case 16:
	case 32:
		color_bit = 3;
		biComp = BI_BITFIELDS;
		break;
	}
	len = (((bmp.bmWidth * bmp.bmBitsPixel) / 8) % 4)
		? ((((bmp.bmWidth * bmp.bmBitsPixel) / 8) / 4) + 1) * 4
		: (bmp.bmWidth * bmp.bmBitsPixel) / 8;

	if((hdc = GetDC(NULL)) == NULL) {
		return NULL;
	}

	hsize = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * color_bit;
	if((ret = malloc(hsize + len * bmp.bmHeight)) == NULL) {
		err = GetLastError();
		ReleaseDC(NULL, hdc);
		SetLastError(err);
		return NULL;
	}
	// DIBヘッダ-- DIB header
	ZeroMemory(ret, hsize);
	((BITMAPINFO *)ret)->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	((BITMAPINFO *)ret)->bmiHeader.biWidth = bmp.bmWidth;
	((BITMAPINFO *)ret)->bmiHeader.biHeight = bmp.bmHeight;
	((BITMAPINFO *)ret)->bmiHeader.biPlanes = 1;
	((BITMAPINFO *)ret)->bmiHeader.biBitCount = bmp.bmBitsPixel;
	((BITMAPINFO *)ret)->bmiHeader.biCompression = biComp;
	((BITMAPINFO *)ret)->bmiHeader.biSizeImage = len * bmp.bmHeight;
	((BITMAPINFO *)ret)->bmiHeader.biClrImportant = 0;

	// DIB取得-- DIB to obtain
	if(GetDIBits(hdc, hbmp, 0, bmp.bmHeight, ret + hsize, (BITMAPINFO *)ret, DIB_RGB_COLORS) == 0) {
		err = GetLastError();
		ReleaseDC(NULL, hdc);
		free(ret);
		SetLastError(err);
		return NULL;
	}
	ReleaseDC(NULL, hdc);
	*size = hsize + len * bmp.bmHeight;
	return ret;
}

/*
* dib_to_bitmap - DIBをビットマップに変換-- Convert a bitmap DIB
*/
static HBITMAP dib_to_bitmap(const BYTE *dib)
{
	HDC hdc;
	HBITMAP ret;
	size_t hsize;
	DWORD err;
	int color_bit;

	if((hdc = GetDC(NULL)) == NULL) {
		return NULL;
	}

	// ヘッダサイズ取得-- Header size acquisition
	if((color_bit = ((BITMAPINFOHEADER *)dib)->biClrUsed) == 0) {
		color_bit = ((BITMAPINFOHEADER *)dib)->biPlanes * ((BITMAPINFOHEADER *)dib)->biBitCount;
		if(color_bit == 1) {
			color_bit = 2;
		}
		else if(color_bit <= 4) {
			color_bit = 16;
		}
		else if(color_bit <= 8) {
			color_bit = 256;
		}
		else if(color_bit <= 16) {
			color_bit = 3;
		}
		else if(color_bit <= 24) {
			color_bit = 0;
		}
		else {
			color_bit = 3;
		}
	}
	hsize = sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * color_bit;

	// ビットマップに変換-- Convert to bitmap
	ret = CreateDIBitmap(hdc, (BITMAPINFOHEADER *)dib, CBM_INIT, dib + hsize, (BITMAPINFO *)dib, DIB_RGB_COLORS);
	if(ret == NULL) {
		err = GetLastError();
		ReleaseDC(NULL, hdc);
		SetLastError(err);
		return NULL;
	}
	ReleaseDC(NULL, hdc);
	return ret;
}

/*
* clipboard_copy_data - クリップボードデータのコピーを作成--  Create a copy of the clipboard data
*/
HANDLE clipboard_copy_data(const UINT format, const HANDLE data, size_t *ret_size)
{
	HANDLE ret = NULL;
	BYTE *from_mem, *to_mem;
	LOGPALETTE *lpal;
	WORD pcnt;

	if(data == NULL) {
		return NULL;
	}

	switch(format) {
	case CF_PALETTE:
		// パレット--Palette
		pcnt = 0;
		if(GetObject(data, sizeof(WORD), &pcnt) == 0) {
			return NULL;
		}
		if((lpal = malloc(sizeof(LOGPALETTE) + (sizeof(PALETTEENTRY) * pcnt))) == NULL) {
			return NULL;
		}
		lpal->palVersion = 0x300;
		lpal->palNumEntries = pcnt;
		if(GetPaletteEntries(data, 0, pcnt, lpal->palPalEntry) == 0) {
			free(lpal);
			return NULL;
		}

		ret = CreatePalette(lpal);
		*ret_size = sizeof(LOGPALETTE) + (sizeof(PALETTEENTRY) * pcnt);

		free(lpal);
		break;

	case CF_DSPBITMAP:
	case CF_BITMAP:
		// ビットマップ--Bitmap
		if((to_mem = bitmap_to_dib(data, ret_size)) == NULL) {
			return NULL;
		}
		ret = dib_to_bitmap(to_mem);
		free(to_mem);
		break;

	case CF_OWNERDISPLAY:
		*ret_size = 0;
		break;

	case CF_DSPMETAFILEPICT:
	case CF_METAFILEPICT:
		// コピー元ロック-- Source rock
		if((from_mem = GlobalLock(data)) == NULL) {
			return NULL;
		}
		// メタファイル-- Metafile
		if((ret = GlobalAlloc(GHND, sizeof(METAFILEPICT))) == NULL) {
			GlobalUnlock(data);
			return NULL;
		}
		// コピー先ロック-- Destination lock
		if((to_mem = GlobalLock(ret)) == NULL) {
			GlobalFree(ret);
			GlobalUnlock(data);
			return NULL;
		}
		CopyMemory(to_mem, from_mem, sizeof(METAFILEPICT));
		if((((METAFILEPICT *)to_mem)->hMF = CopyMetaFile(((METAFILEPICT *)from_mem)->hMF, NULL)) != NULL) {
			*ret_size = sizeof(METAFILEPICT) + GetMetaFileBitsEx(((METAFILEPICT *)to_mem)->hMF, 0, NULL);
		}
		// ロック解除-- Unlock
		GlobalUnlock(ret);
		GlobalUnlock(data);
		break;

	case CF_DSPENHMETAFILE:
	case CF_ENHMETAFILE:
		// 拡張メタファイル-- Enhanced Metafiles
		if((ret = CopyEnhMetaFile(data, NULL)) != NULL) {
			*ret_size = GetEnhMetaFileBits(ret, 0, NULL);
		}
		break;

	default:
		// その他-- Other
		// メモリチェック-- Memory check
		//if(IsBadReadPtr(data, 1) == TRUE) {
		//	return NULL;
		//}
		// サイズ取得-- Size acquisition
		if((*ret_size = GlobalSize(data)) == 0) {
			return NULL;
		}
		// コピー元ロック-- Source rock
		if((from_mem = GlobalLock(data)) == NULL) {
			return NULL;
		}

		// コピー先確保-- Copy to ensure
		if((ret = GlobalAlloc(GHND, *ret_size)) == NULL) {
			GlobalUnlock(data);
			return NULL;
		}
		// コピー先ロック-- Destination lock
		if((to_mem = GlobalLock(ret)) == NULL) {
			GlobalFree(ret);
			GlobalUnlock(data);
			return NULL;
		}

		// コピー-- Copy
		CopyMemory(to_mem, from_mem, *ret_size);

		// ロック解除-- Unlock
		GlobalUnlock(ret);
		GlobalUnlock(data);
		break;
	}
	return ret;
}

/*
* clipboard_free_data - クリップボード形式毎のメモリの解放 --
Freeing memory for each clipboard format
*/
BOOL clipboard_free_data(const UINT format, HANDLE data)
{
	BOOL ret = FALSE;
	BYTE *mem;

	if(data == NULL) {
		return TRUE;
	}

	switch(format) {
	case CF_PALETTE:
		// パレット-- Palette
		ret = DeleteObject((HGDIOBJ)data);
		break;

	case CF_DSPBITMAP:
	case CF_BITMAP:
		// ビットマップ-- Bitmap
		ret = DeleteObject((HGDIOBJ)data);
		break;

	case CF_OWNERDISPLAY:
		break;

	case CF_DSPMETAFILEPICT:
	case CF_METAFILEPICT:
		// メタファイル-- Metafile
		if((mem = GlobalLock(data)) != NULL) {
			DeleteMetaFile(((METAFILEPICT *)mem)->hMF);
			GlobalUnlock(data);
		}
		if(GlobalFree((HGLOBAL)data) == NULL) {
			ret = TRUE;
		}
		break;

	case CF_DSPENHMETAFILE:
	case CF_ENHMETAFILE:
		// 拡張メタファイル-- Enhanced Metafiles
		ret = DeleteEnhMetaFile((HENHMETAFILE)data);
		break;

	default:
		// その他-- Other
		if(GlobalFree((HGLOBAL)data) == NULL) {
			ret = TRUE;
		}
		break;
	}
	return ret;
}
