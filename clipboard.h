#pragma once

HANDLE clipboard_copy_data(const UINT format, const HANDLE data, size_t *ret_size);
BOOL clipboard_free_data(const UINT format, HANDLE data);
