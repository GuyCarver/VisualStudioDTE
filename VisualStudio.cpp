//----------------------------------------------------------------------
// Copyright (c) 2021, Guy Carver
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without modification,
// are permitted provided that the following conditions are met:
//
//     * Redistributions of source code must retain the above copyright notice,
//       this list of conditions and the following disclaimer.
//
//     * Redistributions in binary form must reproduce the above copyright notice,
//       this list of conditions and the following disclaimer in the documentation
//       and/or other materials provided with the distribution.
//
//     * The name of Guy Carver may not be used to endorse or promote products derived
//       from this software without specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
// ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
// WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
// DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
// ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
// (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
// LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
// ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
// SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
//
// FILE    VisualStudio.cpp
// DATE    08/24/2021 10:17 PM
//----------------------------------------------------------------------

#include "VisualStudio.h"

//#include <mmsystem.h>
#include <iostream>
#undef calloc
#undef free
#undef malloc
#undef realloc
#undef _expand
#undef _msize

#include <atlbase.h>
#include <stdio.h>

//
// Local data.
//
namespace
{

_DTEPtr spDTE = nullptr;						// Local DTE pointer set by Open() and released by Close()
OutputWindowPanePtr spOW[4];
uint32_t inumPanes = 0;

}	//namespace

//
// <summary> Open a Development Tools Extensibility (DTE) object from a
//  running DevStudio .NET application.  </summary>
/// <param name="apVersion"> The version string to open. </param>
// <returns> True if open succeeded. </returns>
// <remarks> Close() should be called once the DTE is no longer needed.  But be
//  sure to set any COM Interface smart pointers to nullptr before calling Close(). </remarks>
//
bool Open( const wchar_t *apVersion )
{
	//If not already open.
	if (spDTE == nullptr) {
		CoInitialize(nullptr);

		if (apVersion == nullptr) {
			apVersion = L"VisualStudio.DTE.16.0";
		}

		CLSID clsidDTE;
		//Get the CLSID for the DTE from the name of the object.
		HRESULT hr = CLSIDFromProgID(apVersion, &clsidDTE);
		if (SUCCEEDED(hr)) {
			CComPtr<IUnknown> pUnk = nullptr;

			//Get an active DTE object.
			hr = GetActiveObject(clsidDTE, nullptr, (IUnknown**)&pUnk);

			if ((SUCCEEDED(hr)) && (pUnk)) {
				pUnk->QueryInterface(&spDTE);		//Get the interface from the object.
			}
		}
	}

	return(spDTE != nullptr);
}

//
// <summary> Close the DTE object. </summary>
// <remarks> Be sure to set any COM interface smart pointers to nullptr before
//  calling this method or you may crash. </remarks>
//
void Close()
{
	spDTE = nullptr;
}

//
// <summary> Open the given file and set the current line. </summary>
// <param name=apFileName> Name of file so set including full path. </param>
// <param name=aiLineNumber> Line number to set cursor to in file. </param>
// <returns> true = Success. </returns>
//
bool SetFile( const wchar_t *apFileName, uint32_t aiLineNumber )
{
	bool bres = false;

	//If DTE open.
	if (spDTE) {
		wchar_t ptemp[MAX_PATH];
		swprintf_s(ptemp, MAX_PATH, L"\"%s\"", apFileName);	//Put quotes around file name in case there are any
												// spaces in the path name.

		//Create BSTR for passing to .NET.
		CComBSTR wparam(ptemp);					//Turn into a BSTR.
		CComBSTR wcommand(L"File.OpenFile");	//OpenFile command.

		//Execute the command.
		auto hr = spDTE->ExecuteCommand(wcommand, wparam);

		//If successful then goto the line number.
		if (SUCCEEDED(hr)) {
			char pparam[64];
			sprintf_s(pparam, 64, "%d", aiLineNumber);	//Turn line number into a string.

			wparam = pparam;					//Assign string to BSTR.
			wcommand = L"Edit.Goto";			//Goto command.

			//Execute the command.
			hr = spDTE->ExecuteCommand(wcommand, wparam);
			bres = (SUCCEEDED(hr));
		}
	}
	return(bres);
}

///
/// <summary> Get the breakpoints object and number of breakpoints. </summary>
/// <param name="aCount"> OUTPARAM for the # of breakpoints. </param>
/// <returns> Pointer to Breakpoints object. </returns>
///
void *GetBreakPoints( uint32_t &aCount )
{
	aCount = 0;
	Breakpoints *pbreaks = nullptr;

	if (spDTE) {
		DebuggerPtr pdeb;
		auto hr = spDTE->get_Debugger(&pdeb);
		if (SUCCEEDED(hr)) {
			hr = pdeb->get_Breakpoints(&pbreaks);
			if (SUCCEEDED(hr)) {
				long c = 0;
				pbreaks->get_Count(&c);
				aCount = c;
			}
		}
	}

	return pbreaks;
}

///
/// <summary> Get a breakpoint data. </summary>
/// <param name="apBreaks"> Pointer to the Breakpoints object. </param>
/// <param name="aIndex"> 0 based index for the break point. </param>
/// <returns> BreakPointData structure </returns>
///
const BreakPointData GetBreakPoint( void *apBreaks, uint32_t aIndex )
{
	BreakPointData res;

	if (apBreaks) {
		auto *pbreaks = reinterpret_cast<Breakpoints*>(apBreaks);
		VARIANT v;
		v.llVal = aIndex + 1;					// Add 1 because the index needs to be 1 based instead of 0.
		v.vt = VT_UINT;
		BreakpointPtr pbreak;
		auto hr = pbreaks->Item(v, &pbreak);
		if (SUCCEEDED(hr)) {
			BSTR f;
			pbreak->get_File(&f);
			res.FileName = f;
			long ln;
			pbreak->get_FileLine(&ln);
			res.Line = ln;
			VARIANT_BOOL en;
			pbreak->get_Enabled(&en);
			res.bEnabled = en != 0;
		}
	}

	return res;
}

//
// <summary> Execute a NET command. </summary>
// <param name=apCommand> Command string to execute in DevStudio. </param>
// <param name=apArgs> Argument string to send with command. </param>
// <returns> true = Success. </returns>
//
bool SendCommand( const wchar_t *apCommand, const wchar_t *apArgs )
{
	bool bres = false;

	//If opened ok.
	if (spDTE) {
		//TODO: See if it's better for params to be wchar_t.
		//Turn strings into BSTR for passing to .NET.
		CComBSTR wcommand(apCommand);
		CComBSTR wparam(apArgs);

		//Send the command and arguments.
		auto hr = spDTE->ExecuteCommand(wcommand, wparam);

		bres = (SUCCEEDED(hr));
	}
	return(bres);
}

//
// <summary>  </summary>
//
uint32_t AddOutputWindow( const char *apWindowName )
{
	uint32_t iindex = 0;

	HRESULT hr;

	//If opened ok.
	if (spDTE) {
		WindowsPtr pwin = nullptr;

		hr = spDTE->get_Windows(&pwin);

		WindowPtr pwin2 = nullptr;
		VARIANT v;
		memset(&v, 0, sizeof(v));
		v.vt = VT_I4;
		
		v.lVal = vsWindowTypeOutput;

		hr = pwin->Item(v, &pwin2);

		if (pwin2) {
			OutputWindowPtr spout = nullptr;
			CComPtr<IDispatch> pdisp = nullptr;
		
			vsWindowType etype;

			hr = pwin2->get_Type(&etype);
			
			hr = pwin2->get_Object(&pdisp);
			if (pdisp) {
				hr = pdisp->QueryInterface(&spout);
			}

			if (spout) {
				OutputWindowPanesPtr ppanes = nullptr;

				hr = spout->get_OutputWindowPanes(&ppanes);
				if (ppanes)
				{
					OutputWindowPanePtr spowp = nullptr;
					CComBSTR wname(apWindowName);
						
					hr = ppanes->Add(wname, &spowp);
				
					if (spowp)
					{
						spOW[inumPanes] = spowp;
						iindex = inumPanes++;
					}
				}
			}
		}
	}

	return(iindex);
}

//
// <summary>  </summary>
//
void OutputToPane( uint32_t aiIndex, const char *apString )
{
	
}
