# Third-Party Licenses

OpenSpace는 다음 오픈소스 소프트웨어에 의존합니다.
각 라이브러리의 원본 라이선스 전문은 해당 프로젝트 저장소에서 확인할 수 있습니다.

---

## Python 패키지

| 패키지 | 라이선스 | 프로젝트 URL |
|--------|---------|-------------|
| FastAPI | MIT | https://github.com/tiangolo/fastapi |
| Uvicorn | BSD-3-Clause | https://github.com/encode/uvicorn |
| SQLAlchemy | MIT | https://github.com/sqlalchemy/sqlalchemy |
| pydantic-settings | MIT | https://github.com/pydantic/pydantic-settings |
| argon2-cffi | MIT | https://github.com/hynek/argon2-cffi |
| cryptography | Apache-2.0 OR BSD-3-Clause | https://github.com/pyca/cryptography |
| openpyxl | MIT | https://foss.heptapod.net/openpyxl/openpyxl |
| python-multipart | Apache-2.0 | https://github.com/Kludex/python-multipart |
| Jinja2 | BSD-3-Clause | https://github.com/pallets/jinja |
| aiofiles | Apache-2.0 | https://github.com/Tinche/aiofiles |

## 프론트엔드 라이브러리 (로컬 번들: `web/static/lib/`)

모든 프론트엔드 라이브러리는 CDN이 아닌 로컬 번들로 포함되어 있어 폐쇄망에서도 동작합니다.

| 라이브러리 | 버전 | 로컬 파일 | 라이선스 | 프로젝트 URL |
|-----------|------|----------|---------|-------------|
| Jspreadsheet CE | v4 | `jspreadsheet.css`, `jspreadsheet.js` | MIT | https://github.com/jspreadsheet/ce |
| jSuites | latest | `jsuites.css`, `jsuites.js` | MIT | https://github.com/jsuites/jsuites |
| @jspreadsheet/formula | v2 | `jspreadsheet-formula.js` | MIT | https://github.com/jspreadsheet/formula |

---

## 라이선스 전문 요약

### MIT License

```
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

### BSD 3-Clause License

```
Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
   this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.
3. Neither the name of the copyright holder nor the names of its contributors
   may be used to endorse or promote products derived from this software
   without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED.
```

### Apache License 2.0

```
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
