import traceback
import xlwings as xw
import win32api
import win32con


def msgbox_yesno(xb, title, body):  # yes : 6 / no : 7
    return True if win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_YESNO) == 6 else False


def msgbox_ok(xb, title, body):  # ok : 1
    return True if win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_OK) == 1 else False


def msgbox_yesnocancel(xb, title, body):  # yes : 6 / no : 7 / cancel : 2
    ret = win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_YESNOCANCEL)
    if ret == 2:
        exit(0)
    return True if ret == 6 else False


def msgbox_okcancel(xb, title, body):  # ok : 1 / cancel : 2
    ret = win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_OKCANCEL)
    if ret == 2:
        exit(0)
    return True


def msgbox_error_stacktrace(xb, title):  # ok : 1
    return True if win32api.MessageBox(xb.app.hwnd, str(traceback.format_exc()), title, win32con.MB_ICONERROR) else False


def msgbox_error_stacktrace_and_exit(xb, title):  # ok : 1
    win32api.MessageBox(xb.app.hwnd, str(traceback.format_exc()), title, win32con.MB_ICONERROR)
    exit(0)


def msgbox_info(xb, title, body):  # ok : 1
    return True if win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_ICONINFORMATION) else False


def msgbox_warning(xb, title, body):  # ok : 1
    return True if win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_ICONWARNING) else False


def msgbox_retrycancel(xb, title, body):  # retry : 4 / cancel : 2
    return True if win32api.MessageBox(xb.app.hwnd, body, title, win32con.MB_RETRYCANCEL)==4 else False


if __name__ == '__main__':
    msgbox_retrycancel(xw.Book('TC_HELPER.xlsx'), 'title', 'test')
