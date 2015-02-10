#ifndef CAS_PROCESSOR_H
#define CAS_PROCESSOR_H


MIDL_INTERFACE("5152E6FD-AF2A-4c4f-8348-C5A1AF5EE919") ICasProcessor : public IUnknown
{
	STDMETHOD(SetSpecificServiceDecoding)(BOOL fSpecificService) = 0;
	STDMETHOD(GetSpecificServiceDecoding)(BOOL *pfSpecificService) = 0;
	STDMETHOD(SetEnableContract)(BOOL fEnable) = 0;
	STDMETHOD(GetEnableContract)(BOOL *pfEnable) = 0;
	STDMETHOD(GetCardReaderName)(BSTR *pName) = 0;
	STDMETHOD(GetCardID)(BSTR *pID) = 0;
	STDMETHOD(GetCardVersion)(BSTR *pVersion) = 0;
	STDMETHOD(GetInstructionName)(int Instruction, BSTR *pName) = 0;
	STDMETHOD(GetAvailableInstructions)(UINT *pAvailableInstructions) = 0;
	STDMETHOD(SetInstruction)(int Instruction) = 0;
	STDMETHOD(GetInstruction)(int *pInstruction) = 0;
	STDMETHOD(BenchmarkTest)(int Instruction, DWORD Round, DWORD *pTime) = 0;
};


#endif
