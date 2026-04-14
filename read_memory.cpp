/*
Continually reads from freecell.exe's memory and sends the data to a server running on localhost:5000
The data sent includes the current state of the cards and the finished piles.

g++ .\read_memory.cpp -o find_cards.exe -lws2_32
*/

#include <winsock2.h>
#include <windows.h>
#include <tlhelp32.h>
#include <iostream>

#pragma comment(lib, "ws2_32.lib")

#define CARD_OFFSET 0x00008ab0
#define FINISH_OFFSET 0x00008100

bool SendAll(SOCKET sock, const char* data, int len)
{
    int total = 0;
    while (total < len)
    {
        int sent = send(sock, data + total, len - total, 0);
        if (sent == SOCKET_ERROR || sent == 0)
            return false;
        total += sent;
    }
    return true;
}

/**
 * Function wich find the process id of the specified process.
 * \param lpProcessName : name of the target process.
 * \return : the process id if the process is found else -1.
 */
DWORD GetProcessByName(const char* lpProcessName)
{
    PROCESSENTRY32 ProcList{};
    ProcList.dwSize = sizeof(ProcList);

    const HANDLE hProcList = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hProcList == INVALID_HANDLE_VALUE)
        return -1;

    if (!Process32First(hProcList, &ProcList))
        return -1;

    do {
        if (lstrcmpiA(ProcList.szExeFile, lpProcessName) == 0)
            return ProcList.th32ProcessID;

    } while (Process32Next(hProcList, &ProcList));

    return -1;
}

/*
Get base address (0x1000000) of the specified module in the specified process.
*/
uintptr_t GetModuleBaseAddress(DWORD processId, const char* moduleName)
{
    MODULEENTRY32 modEntry{};
    modEntry.dwSize = sizeof(modEntry);

    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32, processId);
    if (hSnapshot == INVALID_HANDLE_VALUE)
        return 0;

    if (!Module32First(hSnapshot, &modEntry))
    {
        CloseHandle(hSnapshot);
        return 0;
    }

    do {
        if (lstrcmpiA(modEntry.szModule, moduleName) == 0)
        {
            uintptr_t base = reinterpret_cast<uintptr_t>(modEntry.modBaseAddr);
            CloseHandle(hSnapshot);
            return base;
        }
    } while (Module32Next(hSnapshot, &modEntry));

    CloseHandle(hSnapshot);
    return 0;
}

int main()
{
    const char* process = "freecell.exe";

    DWORD pid = GetProcessByName(process);
    std::cout << "Freecell PID: " << pid << "\n";
    if (pid == -1)
    {
        std::cout << "Process not found\n";
        return 1;
    }

    uintptr_t baseAddress = GetModuleBaseAddress(pid, process);
    if (baseAddress == 0)
    {
        std::cout << "Failed to locate module base\n";
        return 1;
    }

    uintptr_t cardAddress = baseAddress + CARD_OFFSET;
    uintptr_t finishAddress = baseAddress + FINISH_OFFSET;

    std::cout << "Base: 0x" << std::hex << baseAddress
              << " Cards: 0x" << cardAddress
              << " Foundations: 0x" << finishAddress << std::dec << "\n";

    HANDLE hProcess = OpenProcess(PROCESS_VM_READ, FALSE, pid);
    if (!hProcess)
    {
        std::cout << "Failed to open process\n";
        return 1;
    }

    // initialize sockets
    WSADATA wsa;
    WSAStartup(MAKEWORD(2,2), &wsa);

    SOCKET sock = socket(AF_INET, SOCK_STREAM, 0);

    sockaddr_in server{};
    server.sin_family = AF_INET;
    server.sin_port = htons(5000);
    server.sin_addr.s_addr = inet_addr("127.0.0.1");

    int result = connect(sock, (sockaddr*)&server, sizeof(server));

    if (result == SOCKET_ERROR)
    {
        printf("Socket connect failed: %d\n", WSAGetLastError());
        return 1;
    }

    int cards[9][21];
    int finished[4];

    SIZE_T bytesRead;

    std::cout << "Start reading memory and sending data to server at 127.0.0.1:5000\n";

    while (true)
    {
        if (!ReadProcessMemory(
            hProcess,
            (LPCVOID)cardAddress,
            cards,
            sizeof(cards),
            &bytesRead) || bytesRead != sizeof(cards))
        {
            std::cout << "ReadProcessMemory cards failed\n";
            break;
        }

        if (!ReadProcessMemory(
            hProcess,
            (LPCVOID)finishAddress,
            finished,
            sizeof(finished),
            &bytesRead) || bytesRead != sizeof(finished))
        {
            std::cout << "ReadProcessMemory foundations failed\n";
            break;
        }

        if (!SendAll(sock, (const char*)cards, sizeof(cards)))
        {
            printf("Socket send failed: %d\n", WSAGetLastError());
            break;
        }

        if (!SendAll(sock, (const char*)finished, sizeof(finished)))
        {
            printf("Socket send failed: %d\n", WSAGetLastError());
            break;
        }

        // Python solver currently consumes one snapshot per run.
        break;
    }

    closesocket(sock);
    WSACleanup();
}