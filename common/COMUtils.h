#pragma once
#include <common.h>

#include <Unknwn.h>

class IUnknownImpl : public IUnknown {
public:
    virtual ~IUnknownImpl() = default;

    // IUnknown
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

private:
    long refCount = 1;
};
