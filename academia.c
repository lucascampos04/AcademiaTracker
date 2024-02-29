#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <windows.h>

int main(){
    SHELLEXECUTEINFO shExecInfo = {0};
    shExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
    shExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
    shExecInfo.hwnd = NULL;
    shExecInfo.lpVerb = NULL;
    shExecInfo.lpFile = "python";
    shExecInfo.lpParameters = "C:\\Users\\Lucas\\Documents\\Academia\\main.py";
    shExecInfo.lpDirectory = NULL;
    shExecInfo.nShow = SW_HIDE;
    
    if (ShellExecuteEx(&shExecInfo)){
        WaitForSingleObject(shExecInfo.hProcess, INFINITE);
        CloseHandle(shExecInfo.hProcess);

        FreeConsole();
    } else {
        MessageBox(NULL, "Erro ao executar o script Python.", "Erro", MB_OK | MB_ICONERROR);
    }

    return 0;
}
