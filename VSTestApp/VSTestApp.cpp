// VSTestApp.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <VisualStudio.h>
#include <chrono>
#include <thread>


int main()
{
	Open(nullptr);
	SetFile(L"VisualStudio.cpp", 40);
	std::chrono::seconds st(5);
	std::this_thread::sleep_for(st);

	SetFile(L"VisualStudio.h", 10);
	std::cout << "Hello World!\n";

	uint32_t count = 0;
	auto pbps = GetBreakPoints(count);
	for ( uint32_t i = 0; i < count; ++i) {
		BreakPointData data = GetBreakPoint(pbps, i);
		std::wcout << data.FileName << data.Line << data.bEnabled << std::endl;
	}
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
