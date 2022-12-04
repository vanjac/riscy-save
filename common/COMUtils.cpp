#include "COMUtils.h"

STDMETHODIMP_(ULONG) IUnknownImpl::AddRef() {
    return InterlockedIncrement(&refCount);
}

STDMETHODIMP_(ULONG) IUnknownImpl::Release() {
    long r = InterlockedDecrement(&refCount);
    if (r == 0) {
        delete this;
    }
    return r;
}
