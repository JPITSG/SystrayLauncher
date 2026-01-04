#pragma once
#include <guiddef.h>
typedef struct _tagWINEVTOKEN {
ULONGLONG value;
} WINRT_IMAGING_EVENT_TOKEN;
typedef WINRT_IMAGING_EVENT_TOKEN EventToken;



typedef struct EventRegistrationToken {
  __int64 value;
} EventRegistrationToken;
