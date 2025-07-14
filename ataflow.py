import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageDraw, ImageFont
import io
import base64
import csv

icon_base64 = """
iVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAEyWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSfvu78nIGlkPSdXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQnPz4KPHg6eG1wbWV0YSB4bWxuczp4PSdhZG9iZTpuczptZXRhLyc+CjxyZGY6UkRGIHhtbG5zOnJkZj0naHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyc+CgogPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9JycKICB4bWxuczpBdHRyaWI9J2h0dHA6Ly9ucy5hdHRyaWJ1dGlvbi5jb20vYWRzLzEuMC8nPgogIDxBdHRyaWI6QWRzPgogICA8cmRmOlNlcT4KICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0nUmVzb3VyY2UnPgogICAgIDxBdHRyaWI6Q3JlYXRlZD4yMDI1LTA1LTI5PC9BdHRyaWI6Q3JlYXRlZD4KICAgICA8QXR0cmliOkV4dElkPmJlY2FkMDFkLTgyN2EtNGFiYS05MGYzLTY3MWVkMDk2N2JhNDwvQXR0cmliOkV4dElkPgogICAgIDxBdHRyaWI6RmJJZD41MjUyNjU5MTQxNzk1ODA8L0F0dHJpYjpGYklkPgogICAgIDxBdHRyaWI6VG91Y2hUeXBlPjI8L0F0dHJpYjpUb3VjaFR5cGU+CiAgICA8L3JkZjpsaT4KICAgPC9yZGY6U2VxPgogIDwvQXR0cmliOkFkcz4KIDwvcmRmOkRlc2NyaXB0aW9uPgoKIDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PScnCiAgeG1sbnM6ZGM9J2h0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvJz4KICA8ZGM6dGl0bGU+CiAgIDxyZGY6QWx0PgogICAgPHJkZjpsaSB4bWw6bGFuZz0neC1kZWZhdWx0Jz5BVEEgRkxPVyAtIDE8L3JkZjpsaT4KICAgPC9yZGY6QWx0PgogIDwvZGM6dGl0bGU+CiA8L3JkZjpEZXNjcmlwdGlvbj4KCiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0nJwogIHhtbG5zOnBkZj0naHR0cDovL25zLmFkb2JlLmNvbS9wZGYvMS4zLyc+CiAgPHBkZjpBdXRob3I+TGFyaXNzYSBLZXRsaW48L3BkZjpBdXRob3I+CiA8L3JkZjpEZXNjcmlwdGlvbj4KCiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0nJwogIHhtbG5zOnhtcD0naHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyc+CiAgPHhtcDpDcmVhdG9yVG9vbD5DYW52YSAoUmVuZGVyZXIpIGRvYz1EQUdaRmlESEVoNCB1c2VyPVVBRXJVdzh4SDZRIGJyYW5kPUxhcmlzc2FrZXRsaW5iYXRpc3RhQGdtYWlsLmNvbSB0ZW1wbGF0ZT08L3htcDpDcmVhdG9yVG9vbD4KIDwvcmRmOkRlc2NyaXB0aW9uPgo8L3JkZjpSREY+CjwveDp4bXBtZXRhPgo8P3hwYWNrZXQgZW5kPSdyJz8+9HW0hgAAFthJREFUeJzt3XuUHFWBBvCvJpOQVxUgJMCWrrUImkkikQPConaldmUzkV2I4IIoPgZldTmK7uIeD2yXC+zp9rGrq3L0qERhVHQXEXkeYsZ9VKrXPUoWfJGHa4SKnAqQwQSqk8lrMrV/VHXS6enuutV9J9PJ/X7nNJPpvn3rTnXV11V1b100EJGytOluABFNHwYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKUzD27/xteluBBFNj34AH5juRhDR9Oib7gYQ0fRhABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHC+iXXF0uuj4imkLwAiOMYwAEAu6XVSURTSl4AaIgR41H8dMbl2BpIq5aIpo7sUwBg6zCAddKrJSL5puAiIHd+omMFewGIFMYAIFIYA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEAiBTGACBSGAOASGEMACKFMQCIFMYAIFIYA4BIYfInBJFp4QAwsFCb9LymAXGc/Vy9XTrweDWWN1/BcuD1czXMHRNbfpatW2MEgZSWAQCWL5+83mTpZP1P1fLH5gLrx+R+rufpGuZXxZbf7rmpInH5vRsAi68Elqy00YczAQnbcjwxDn3LQ/DWRV3XtWAxsPQNZ2DBWYPQJO1np038HHPX/hwb7+2+roXXAwvPexv6tPlS1l0vi2Ng3m9GsGF0G0Y3dl+f8wYdC1+1Clpf7+4bEvXuH7nkSg2IrkOMayRtw1UsWPITABICYAmwYPEiYN83um8WgBgxNPwjLl8mJwBe8YwG7bxPIcaZ0I73iZpjYOGSFdj+WzkBsHDxadD2fQmI9e4r6329GwAAAA3QoKFnv8Y0QMZ1FC2eAADMHO+6qsaa0bPrThZtYgrqBBS5PqbEH0lEzTEAiBTGACBSGAOASGEMACKFMQCIFMYAIFIYA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEAiBTGACBSGAOASGE9PiEITb+49p9Y3twicVqRrPnUetWhSfpkTsuUTvIiZ9UxAKi9WIuB+DkcnHE9+vq735BnjgPj438JTbsGx/9sRUAce+ifcTtO3tn937p9AaDt+ziAC9OZsrrGACARY3h2w8N4cWv3028NXqRhm36OhDYdA2IACDEx+0FseChGHGcf9NTKNPs5ugBYuvLdacUMADqKtm8Atqzpvp6rlsTYpsR8m0j20Rg4OAPwPDlVLl0pp54ULwISKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARArjhCAkoh+Lr3gFzn1n91OCbQCg7U1mBDnOJwQ7FjAAqD0tBoCXY7b2BLBPVqWzwd2/JzAAKIOmAeiDFp/Mffb4wwAgEdKmoabewouARApjABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApjAFApDDOB0BZYiBOf/akXm3XMYEBQO3FMQDsBrAGQNyD84L0YfuCFzC6YLrbcUxiAFAGDUD8PHbuvxonnhInM4T1kLEXAHw6xqg33S05JjEASMzTDwNb1vBw+zjDi4BECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcLkTwhy9Xc1HJRVbZxOQCFxFporvyepsn1qTY6x9O3Audf22HRAAPQqcOfvYuC26W5JE3VzKUrb7iI51aTkBUAyW9xFmBj7oaT9NQawFDFk7f9zoe27E9i3R0ptwMkANKnh1JNiADgDs8bXyN74pNgFDdeecROevOQJrH90ultzpGTXfzO0aA3kbSjLJNYlMQA0DYjjhdC0i6XVGceaxP2rH5r2Rmm1JZPlHu97P6BpGhDPASR+rlLFMXbPOwm75013QybTAMQ4A5p2urQ65e4TUk8BtHTCSHnNkzkBpaZpkPp1rcCX/yGy151M2sR0t6A1rbaz9uY+AV4EJFIaA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEAiBTGACBSGAOASGEMACKFMQCIFMYAIFIYA4BIYQwAIoX1ozZvCR2v+Pl2pjfXWxxD5rwb/UC8U0ZF1IVYi6HFe7F9gaz6ACAC4h3o2Yk8pIkBjEusbQKIX4SGg9LqlCVGDEBHHM+UNTFIP06pni2lJupCHOPJ0b2451NyqnsiirFs/M2Ys+f4P8XbMwd4ZH2E0Xvl1OdVArzl/HMxf3fvBefoqYC275sALoG0I4CND+449Nvhw4vmsl7vtg6Vl79pe/t6c1kHbJj7IuY2mf+0V//+Tt87NgcYHWvfpjxGfz2BDdt2Qt+Vv30iZbr5+7efCixdeaB9A/LpvZQjotau+voPAKyCpkk5ujv+DxGJqCUGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHCGABECmMAEB0zbpFeYz+WPPF+6bUSkXzLPeAFvBISJ/LRcNXXe2/yQyJqTqvt/HJmBe2XNbUQER17uPMTKYwBQKQwBgCRwhgARApjABApjAFApDAGAJHCGABECmMAECmMAUCkMAYAkcIYAEQKYwAQKYwBQKQwBgCRwhgARApjABApTMOSy1493Y0goukhbXJBIjr2MACIFMYAIFIYA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEA6LhWdMyTAZwMQE8fswGMAagC2AXg92UvjKa4DTMAnJA+ZpS98IWpXF4euQOg6Jh/Zlv6m2zLmA9gPoB5AOYCmJP+nAVgJoAZACaQrOStJS+8teyFT+VYzjLXMT+U1tOf/pzR5Pe+hn83PrQmf2dc95gAMF7ywitFP5iCpb/adcx32JYxr24dzEvXwRwkG1l/uh5iAHsAPOcH0VcHhzf/SHQdiNhz6wVXALikw7drSNbbTCTtrT0af48B7Cx54Y1lL3yqYOnayNDAHQAWAjgRyd88C4fXd5zWHwOYKHnhZ8peeE+HbRRWdMyTAFzqOuZFAJYCGABwasbbYgDPA/g1gI0lL1wHYE2eUNhz6wUfBvBXSD73E5o8ZtQV37tieNOCSlDd1aT9F7mOeTuOXPcz6n/3g+i2weHNX2vWjqJjXu465m04vE/0ofn+sW3F8KbXV4Iq+kX/SAAoWPoM1zHvAGDleR8AuI75UtkLb8jxljciWalHw34/iHaKFh4ZGvgEgHflXYhtGWcDOCfv+zK8AcD7JdfZyo0AYFuGAeC6HO+bPzXNAQqWrrmOeaVtGe8H4CAJojw0AKenj+WuY14PYL/rmL4fRLeXvPDhSlBtW0HJC3/hOqbo5zrbtoyVlaD6/cYXXMd8G4Dz273ZtowPAmgaAK5jvgfAa7Ma4AfR6trflOtuQNcxL0cHO3/qPQVL13Msy+xwOZ34XSWoHhQpWHTMPwBwVYfLeW3RMZ0O39vKTMn1tXLAD6Kt6b9PzPne3bIbAwBrhxatHBkaeNy2jHsArED+nb+VWQAuti3joZGhgfVrhxa9uV1hP4j+G8A20cpdx7ysxUutnq93btExlzU+WbD0eUjWQZaDflC9q/ZLrgCwLePGPOUbGK5jvjdH+aMZAJtEC9qW/iF0saG5jpnnKEjE0QqAbXUhKRzkqUmHu90oWPqsPbdesNq2jDUAzpVZdxPn25YxsnZo0acKlt70lLkSVGMA9+Wo85KCpdefFqDomIsAnC3y5mb7kW0Zf47kFDzLj8peWAty8QAoOuaFAC4SLd+MbRkfLljC287p3SwrDz+InhApV7D0OekhWDcuKzrmH3ZZR71cp3FdCOv+nTcApB0BFCx91sjQwIPIdwrSrT7bMm4aGRr4SqsCJS+8O0d9p9iWUah/wrb0v8jx/msKln5E8LuOeYXIG0teuLr+d+EAcB3zb0XLtvEa2zJWCpZdKGF5Qvyg+phIuTR5T+lycf22pV/fZR31jsYRwC4Av6j7fV4H75diZGjgiwBEtyHZPrh2aFHT4Cl74WMAHhetyHXMS+t/ty1jVY52LEy/8QEABUufDbELwc/7QfRQ/RNC3x7pN1ZWwhzEkVc7m3Id84ayF/4wq1zJC28BYLR63bb0VbZlXJlVDwD4QXSdH1T3tnn9v7LqKFi6ZlvG32QUE1oHtmVcV7D02ypt2pRDngDYUvLCdyFpZ+0x3vCz8bmxshc2tnNaAqDomDaATo7ARtJv6CfTtrzMtvRltmW8G8Cb8lRkW8Y/FSz9nkpQnXRl0A+ir9iW8XXBqlYB+BgAFB3zFAB/nKcdrmMOlb3wgbRNgxA4KvODaLgSVMfrnxMKgPS8td2Gth/A3QDeJ1DdyqJjnl32wt+0K1T2wofbtunWCxYLLAsA4AfVb5e9cL9o+WZsy7gEwGsyit0JsZ6LU13HvGZwePM3umlTKk8APFf2wp9KWKbIuWY9KQHgOqaLfF3XB/0gum5wePNw4wtl4KcA7lg7tOijtmV8IUedJ7uO+a7B4c2TTgdKXvivI0PGZwGcJFDPq4qOubTshU8i+fbOeyr3lqJjLix74XbBw//YD6qTtrfMU4CCpc9H9vnWf5TE+3n70v79bomeh074QdTVzg8InQK9WPLCmwDsE6nPtgxZFwPzBMCYpGXmDYCul1t0zNMAtL0a38gPos812/nrDQ5v/iKAf89Tr20ZlzZ7vhJUxwB8U7we/VKgba9AO7NsS69dC2jangbrmn3pZqaO65jvQ0ai+UF0jx9EHmDuglif77UFS/9Es8OoHEQ3wn1Z/bhZ0m6XP80o9kDZC3e4jlkBcLFAtcuKjrm8nAw86Uaeb46XFZO+5voBVLWfWvrvcQCjfhA9VgmqrQZG5Q2APTnLN7Mc+XqtJvyg+i8iBUte+H3XMUU+s5qWfe0lL/ya65gfgcCRim0Zby1Y+mch1n3X7P3X2lZ1M5KRjm35QdT01KTtxlOw9D7bMj6aUffukhfeVwmq+wH8CMDlWY1B2iU4OLz5SwJlWxHdCLve+NJv/7YfaO0qsB9ED9mWIbQxpddDug2APEcA57uOOWkASjMlD2+tBNUHW7yc5xpA7AdR15+Bbeln5XzLU2UvfF6wbJCz7pe1eqHshZtcx1yHZFBSlvNHhgauQZtrXRle6zrmLQLldpS8sGk3ZdtEdR3zrQDOzKj8/tqwxpIXPpRR9hDbMm5o1a8qaI5gua4OP4uOeQaAqzOKPeMH0X8CgB9UH8HhobBZVhUd85XdtA9T1wvwdKsXbEvPcwSwN+0n74ptGZnfcg1+n6PsgZx1t/17/CBq2V3YoA/Ap3Muu9GFAmXubnXBuW0AiAz8KXnhXXW/PoLkEFLEq3N0CTYzW7BcV33Q6cCfE9qV8YPortpGXvbCp5FcbRYho0twKsYBHPCDaHOrF23LyDMOQNZ1B6GRmnVEvyAAYEHOutuGS8kL7wfwnGBdp7V5bQOAbm9Uikte2LJnomUApAN/3phR+VO1bz4ASG+m+Yloy9JzpU6JBkDHG2DB0ufalvHXGcUm6odWAoAfRPeLLiPtEsyzsTaaiiOA/0tP6VrJEwBSBgGVxA/na85qHCzTiuuYr89Z94Z2L1aC6oFW59x5+EH0PQD3dlnNY2Uv/FWrF1t+e7iOKTLs96SRoYGfNTyXZwjviqJjvqbshb/O8Z6aKT8CEBz4M+465v2uc8SfnWes/Clpt9Lq7KJN5TkC+BmSjXcCh++EjJEcte0DMFbywh0AfplRj+i6B+QdAWS1qdFc1zGvGBze3LZ3Kr0/5T15KvaDKHMcix9U77At4yZ0cYTmB9WH/aBquI7Z8c1efhC13a6aNq7omBayB/4AycWQlhdEBPS5jvnhnHcJ1rQ9LK/TUQAIDvwBkvsCXtfJMmpsy/hIwdJXd9hbIXwEUPLCb5W9ME+fdyui6x6QdATgB5EPmDuQY3uzLeMLRcd8vOyFW5q9XrD02SNDA3cj3ynAmB9Uv5tVqOyFz7iO+SjEbvBp5hk/iNIvV/NpAH/UQR1RVvd80wBID82P1hjz9xYsvVgJqnnPdab0CMC2jMsAHK3/bdpS2zL+pBJUM0ckNpHnc5LRHQfkuxlqrOiYOpLQmIXD98jPavGYieSiXNUPoo2VoDoKHDqs/rJtGZ/IsezTXcd83Lb0L/tBdQ2ALUgGrZ1hW3oh7eHKGtx1BD+IPi86b0TJC7/aYR8/ADxY+0Lwg+hbtmWIXO1v9G/N5h2oN2njKVj6iTh695cDgO465rXpgIw8RL+FOjoEdR3zY528r1OuY36k7IWdBECeawCyAiDPEUDBdcyOLmSVPNxYCaqfP/x7+JmRIeOdAF6VoxrDtoybbcu4uZM2NHis5IUl0cLJqYL5FLJ70iYpeeEjh+upftu2jH9Azgl82l38q5l0EdB1zA+g837JjqR3CebtEpyyU4CiY14AoJBZUK5Li47ZyWFenvPx6TgC6Mb2+l8qQXV3yQtXIV8XnyxPrhjetCrP/RuVoBr7QdR08o4ML9Xfn1L2wt8C+HHOOn5R9sL1WYWOCICCpc+0LUPkynyI5HBqC4DfIukz3grgGQDPIvmA8vStnmVbRp7bIYEpPAUQ/PbfgSPXwVNIBpQ8g2RyiFEAeW72mdHhEOk8s+3ICoCjNQfBpLEIZS/cUPJCG8l6P1oeWDG8ya4EVdGuvUP8oHon8m0HALC2sRfGD6LhfMttf/Gv5ohTANcx3wHg5Rnv2VPywnPKydXiltYOLbrRtozPiTQiXfYNWTcANZiSU4CiY54JgQugJS98d9kLH82o6zzXMf83x+LfV7D0WypBNU9o5RmVJ+uKfOYdjxLEaDFRS9kLN/pBdJ7rmDen91TkvTtR1K/8IPr7weHNj2QXba7shS+4jnkfgGtE35OOI2h87t6RIeN2iI2AHSt54XdElnXoCKBg6bAt4+8E3vODrJ0fAPygeg/yHQVcXHTMgRzlhQKglPMuwHTYb9aFtd+JdAWVvfBx5JhtCOmdZqKFC5beh+k5BTgaF4i3l72w5TyNlaAaDQ5vvrnkhWcD+Ge0GbmY014AD5e88OoVw5te183OX1PyQtGRgQAw5gfRpC/C9CL5A4J13FcJqi+KFDz0QdqWcRqAHwL4HyQDPWqz/s5NHzMB7Cp5odD4/bIXhq5jrgbwdhyeijkC8BKAF9OfOwHs9INopx9UIz+IhL5Zio6pAViHpO96H5IPbW/9v/0g2usH1f3JTUpi0tmKtgP4MpLrII0zH5+Q1r26ElQnROr0g+iTtmV8Nq1jPw6vg9rjiHWAHBuybRn9AO7C4fv4J3Dkvf5HPOcHUSBad4bvAFiP5FpA7UaimmbDZFsNnW1XNhBpSNkLny174ccBfLzomK+zLX2FbRnnAFiMpBen3dFBjOSUbSOAX5a88Ak/iB7t8ia1Zm38seuY9yIZI9M4I3Vt3oUDSLbd9a2OAEteeIebDB0/mJbfj2Sb35M+dgN4acXwpsxuyhr+fwHouJUe1c7D4f8nwHykXYwAIj+IItHJYI9XDAAihTEAiBTGACBSGAOASGEMACKFMQCIFMYAIFIYA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEAiBTGACBSGAOASGEMACKFMQCIFMYAIFIYA4BIYQwAIoUxAIgUxgAgUhgDgEhhDAAihTEAiBTGACBS2P8DaNKDI7ire5MAAAAASUVORK5CYII=
"""

def centralizar_janela(root, largura, altura):
    w, h = root.winfo_screenwidth(), root.winfo_screenheight()
    x, y = (w - largura)//2, (h - altura)//2
    root.geometry(f"{largura}x{altura}+{x}+{y}")
icon_data = base64.b64decode(icon_base64)
Image.MAX_IMAGE_PIXELS = None
A4_WIDTH, A4_HEIGHT = 2480, 3508

# Função para redimensionar a imagem mantendo a proporção
def redimensionar_imagem(img, largura):
    proporcao = largura / img.width
    nova_altura = int(img.height * proporcao)
    return img.resize((largura, nova_altura), Image.LANCZOS)

# Função para adicionar rodapé
def adicionar_rodape(pagina, rodape_texto="Documento gerado via ATA FLOW"):
    draw = ImageDraw.Draw(pagina)
    font_size = 20

    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except IOError:
        font = ImageFont.load_default()

    text_width, text_height = draw.textbbox((0, 0), rodape_texto, font)[2:]
    x = (pagina.width - text_width) / 2
    y = pagina.height - 100
    draw.text((x, y), rodape_texto, fill="black", font=font)
    return pagina

# Função para converter uma lista de imagens em um único PDF
def imgtopdf_from_list(image_pil_objects, header_img=None):
    pdf_pages = []
    current_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
    y_offset = 0
    spacing = 10

    for index, img in enumerate(image_pil_objects):
        img = redimensionar_imagem(img, A4_WIDTH)

        if y_offset == 0 and index != 0 and header_img:
            header_img_resized = redimensionar_imagem(header_img, A4_WIDTH)
            current_page.paste(header_img_resized, (0, 0))
            y_offset += header_img_resized.height

        if y_offset + img.height > A4_HEIGHT:
            current_page = adicionar_rodape(current_page)
            pdf_pages.append(current_page)
            current_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            y_offset = 0

            if index != 0 and header_img:
                header_img_resized = redimensionar_imagem(header_img, A4_WIDTH)
                current_page.paste(header_img_resized, (0, 0))
                y_offset += header_img_resized.height

        current_page.paste(img, (0, y_offset))
        y_offset += img.height + spacing

    current_page = adicionar_rodape(current_page)
    pdf_pages.append(current_page)

    pdf_bytes = io.BytesIO()
    pdf_pages[0].save(pdf_bytes, "PDF", save_all=True, append_images=pdf_pages[1:])
    return pdf_bytes.getvalue()

# Função para selecionar o arquivo PDF
def selecionar_pdf():
    janela_foi_fechada = False
    root = tk.Tk()
    root.withdraw()
    def fechar_janela():
        nonlocal janela_foi_fechada
        janela_foi_fechada = True
        root.destroy()
    image = Image.open(io.BytesIO(icon_data))
    photo = ImageTk.PhotoImage(image)
    root.iconphoto(False, photo)
    centralizar_janela(root, 1000, 600)
    root.protocol("WM_DELETE_WINDOW", fechar_janela)
    messagebox.showinfo("Seleção de Arquivo", "Selecione o arquivo PDF")
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    root.destroy()
    return file_path

def ler_quantidade_paginas(pdf_path):
    doc = fitz.open(pdf_path)
    num_paginas = len(doc)
    doc.close()
    return num_paginas

def encontrar_frase_final_y0(pdf_path):
    keywords = ["Participantes", "Membros", "Members", "Participants"]
    doc = fitz.open(pdf_path)
    frase_y0 = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("dict")
        blocks = text['blocks']

        for block in blocks:
            if 'lines' in block:
                for line in block['lines']:
                    line_text = ''
                    line_position = None

                    for span in line['spans']:
                        if line_text == '':
                            line_text = span['text'].strip()
                            line_position = span['bbox']
                        else:
                            line_text += ' ' + span['text'].strip()

                    for keyword in keywords:
                        if line_text.startswith(keyword):
                            x0, y0, x1, y1 = line_position
                            frase_y0.append((line_text, y0, page_num))
    doc.close()
    return frase_y0

# Função para identificar os tópicos
def extract_topics_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_topics = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if "lines" in block:
                for i, line in enumerate(block["lines"]):
                    text = ""
                    x0, y0, x1, y1 = None, None, None, None

                    for span in line["spans"]:
                        text += span["text"].strip() + " "
                        if x0 is None or span["bbox"][0] < x0:
                            x0 = span["bbox"][0]
                        if y0 is None or span["bbox"][1] < y0:
                            y0 = span["bbox"][1]
                        if x1 is None or span["bbox"][2] > x1:
                            x1 = span["bbox"][2]
                        if y1 is None or span["bbox"][3] > y1:
                            y1 = span["bbox"][3]

                    text = text.strip()

                    if text and text[0].isdigit():
                        if len(text) < 5 and i + 1 < len(block["lines"]):
                            next_line = block["lines"][i + 1]
                            next_text = ""
                            for next_span in next_line["spans"]:
                                next_text += next_span["text"].strip() + " "
                                if next_span["bbox"][2] > x1:
                                    x1 = next_span["bbox"][2]
                                if next_span["bbox"][3] > y1:
                                    y1 = next_span["bbox"][3]
                            text += ' ' + next_text.strip()

                        potential_topic = {
                            "text": text,
                            "x0": x0,
                            "y0": y0,
                            "x1": x1,
                            "y1": y1,
                            "page_num": page_num + 1
                        }

                        all_topics.append(potential_topic)
        # Inicializa filtered_topics para evitar erro
    filtered_topics = []

    if all_topics:
        x0_values = [topic['x0'] for topic in all_topics]
        min_x0 = min(x0_values)
        minor_x0 = [topic for topic in all_topics if topic['x0'] == min_x0]

        if len(all_topics) > 1:
            second_x0 = all_topics[1]['x0']
            second_rule_topics = [topic for topic in all_topics if topic['x0'] == second_x0]
            filtered_topics = minor_x0 + [topic for topic in second_rule_topics if topic not in minor_x0]
        else:
            filtered_topics = minor_x0

        filtered_topics = sorted(filtered_topics, key=lambda x: (x['page_num'], x['y0']))

    doc.close()
    return filtered_topics

def salvar_topicos_em_csv(topicos, caminho_csv="topicos_extraidos.csv"):
    if not topicos:
        print("Nenhum tópico encontrado para salvar.")
        return

    with open(caminho_csv, mode='w', newline='', encoding='utf-8') as arquivo_csv:
        campos = ['text', 'x0', 'y0', 'x1', 'y1', 'page_num']
        escritor = csv.DictWriter(arquivo_csv, fieldnames=campos)
        escritor.writeheader()
        for topico in topicos:
            escritor.writerow(topico)

    print(f"Tópicos salvos em: {caminho_csv}")

# Função para encontrar a posição da imagem do cabeçalho das páginas 2 em diante
def find_image_position(doc, page_num):
    page = doc.load_page(page_num)
    blocks = page.get_text("dict")["blocks"]
    for block in blocks:
        if "image" in block:
            x0, y0, x1, y1 = block["bbox"]
            return x0, y0, x1, y1
    return None

# Função para extrair uma região específica do PDF e retornar como imagem PIL
def extrair_regiao_para_memoria(pdf_path, x0, y0, x1, y1, page_num):
    doc = fitz.open(pdf_path)
    pagina = doc.load_page(page_num)
    retangulo = fitz.Rect(x0, y0, x1, y1)
    matriz = fitz.Matrix(3, 3)
    imagem_pix = pagina.get_pixmap(matrix=matriz, clip=retangulo)

    if imagem_pix.width <= 0 or imagem_pix.height <= 0:
        raise ValueError("Dimensões da imagem são inválidas")

    img_data = imagem_pix.tobytes("png")
    img = Image.open(io.BytesIO(img_data))
    doc.close()
    return img

# Função para extrair o cabeçalho
def extrair_cabecalho_para_memoria(pdf_path, y1):
    x0 = 0
    y0 = 0
    mm_to_points = 72 / 25.4
    x1 = 210 * mm_to_points
    return extrair_regiao_para_memoria(pdf_path, x0, y0, x1, y1, 0)

# Função para processar o PDF e extrair as imagens dos tópicos selecionados
def processar_pdf(caminho_pdf, topicos_selecionados):
    imagens_extraidas = []
    imagem_cabecalho = None

    if caminho_pdf:
        doc = fitz.open(caminho_pdf)
        num_paginas = len(doc)
        topics = extract_topics_from_pdf(caminho_pdf)
        mm_to_points = 72 / 25.4
        x0 = 0
        x1 = 210 * mm_to_points
        y1_folha = 297 * mm_to_points
        y_topico_rodape = 790

        coordenadas_para_extrair = []

        if num_paginas > 1:
            imagem_cabecalho_pos = find_image_position(doc, 1)
            if imagem_cabecalho_pos:
                x0_img, y0_img, x1_img, y1_img = imagem_cabecalho_pos
                img_cabecalho = extrair_regiao_para_memoria(caminho_pdf, 0, 0, 210 * mm_to_points, y1_img, 1)
                imagem_cabecalho = img_cabecalho

        if topics:
            y1_cabecalho = topics[0]['y0']
            img_cabecalho = extrair_cabecalho_para_memoria(caminho_pdf, y1_cabecalho)
            imagens_extraidas.append(img_cabecalho)

        for i, topic in enumerate(topics):
            text = topic['text']
            x0_topic = topic['x0']
            y0 = topic['y0']
            page_num = topic['page_num'] - 1

            if text not in topicos_selecionados:
                continue

            y1 = y1_folha

            if i < len(topics) - 1:
                next_y0 = topics[i + 1]['y0']
                next_page_num = topics[i + 1]['page_num'] - 1

                if next_page_num == page_num:
                    y1 = next_y0
                else:
                    y1 = y_topico_rodape

                coordenadas_para_extrair.append((caminho_pdf, x0, y0, x1, y1, page_num, text))

                while next_page_num != page_num and page_num < num_paginas:
                    page_num += 1

                    image_position = find_image_position(doc, page_num)
                    y0 = image_position[3] if image_position else 0

                    y1 = y1_folha

                    if i < len(topics) - 1 and next_page_num == page_num:
                        next_y0 = topics[i + 1]['y0']
                        y1 = next_y0
                    else:
                        y1 = y_topico_rodape
                    coordenadas_para_extrair.append((caminho_pdf, x0, y0, x1, y1, page_num, text))

            if i == len(topics) - 1:
                next_page_num = page_num + 1

                if page_num == num_paginas - 1:
                    coordenadas_finais = encontrar_frase_final_y0(caminho_pdf)
                    if coordenadas_finais:
                        y1 = coordenadas_finais[-1][1] if coordenadas_finais else y_topico_rodape

                    coordenadas_para_extrair.append((caminho_pdf, x0, y0, x1, y1, page_num, text))
                else:
                    y1 = y_topico_rodape
                    coordenadas_para_extrair.append((caminho_pdf, x0, y0, x1, y1, page_num, text))

                    page_num = page_num + 1

                    while page_num <= num_paginas - 1:
                        image_position = find_image_position(doc, page_num)
                        y0 = image_position[3] if image_position else 0

                        if page_num < num_paginas - 1:
                            y1 = y_topico_rodape
                        else:
                            coordenadas_finais = encontrar_frase_final_y0(caminho_pdf)
                            y1 = coordenadas_finais[-1][1] if coordenadas_finais else y_topico_rodape

                        coordenadas_para_extrair.append((caminho_pdf, x0, y0, x1, y1, page_num, text))
                        page_num = page_num + 1

        for caminho_pdf, x0, y0, x1, y1, page_num, text in coordenadas_para_extrair:
            try:
                img = extrair_regiao_para_memoria(caminho_pdf, x0, y0, x1, y1, page_num)
                imagens_extraidas.append(img)
                print(f"Extraído tópico: {text} página: {page_num + 1}")
            except ValueError as e:
                print(f"Erro ao extrair imagem: {e}")

        pdf_bytes = imgtopdf_from_list(imagens_extraidas, header_img=imagem_cabecalho)

        root = tk.Tk()
        root.withdraw()
        image = Image.open(io.BytesIO(icon_data))
        photo = ImageTk.PhotoImage(image)
        root.iconphoto(False, photo)
        centralizar_janela(root, 1000, 600)

        messagebox.showinfo("Salvar PDF", "Escolha a pasta e o nome do arquivo PDF gerado")

        output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        root.destroy()
        if output_pdf_path:
            with open(output_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            print(f"PDF salvo em: {output_pdf_path}")
        else:
            print("Nenhum caminho de saída fornecido.")

def criar_interface(caminho_pdf):
    if caminho_pdf:
        topics = extract_topics_from_pdf(caminho_pdf)

        if not topics:
            messagebox.showwarning("Aviso", "Nenhum tópico encontrado no PDF.")
            return

        while True:
            janela_foi_fechada = False
            root = tk.Tk()
            root.title("Seleção de Tópicos")
            root.geometry("500x500")
            root.attributes("-topmost", True)

            def fechar_janela():
                nonlocal janela_foi_fechada
                janela_foi_fechada = True
                root.destroy()
        
            image = Image.open(io.BytesIO(icon_data))
            photo = ImageTk.PhotoImage(image)
            root.iconphoto(False, photo)
            centralizar_janela(root, 500, 500)
            root.protocol("WM_DELETE_WINDOW", fechar_janela)

            frame_principal = tk.Frame(root)
            frame_principal.pack(fill='both', expand=True)

            frame_pesquisa = tk.Frame(frame_principal)
            frame_pesquisa.pack(padx=10, pady=10, fill='x')

            campo_pesquisa = tk.Entry(frame_pesquisa)
            campo_pesquisa.pack(side='left', fill='x', expand=True)
            
            # Função que simula o placeholder
            def placeholder_entry(entry, texto):
                entry.insert(0, texto)
                entry.config(fg='grey')

                def on_focus_in(event):
                    if entry.get() == texto:
                        entry.delete(0, tk.END)
                        entry.config(fg='black')

                def on_focus_out(event):
                    if entry.get() == '':
                        entry.insert(0, texto)
                        entry.config(fg='grey')

                entry.bind('<FocusIn>', on_focus_in)
                entry.bind('<FocusOut>', on_focus_out)

            # Ativa o placeholder
            placeholder_entry(campo_pesquisa, "Filtrar tópicos...")

            # Variáveis globais do escopo da função
            topicos_ordenados = topics.copy()  # mantém a ordem original extraída
            var_checkbuttons = []
            checkbox_widgets = []
            selecoes = {}
            # Função que mostra os tópicos
            def mostrar_topicos(topicos_para_mostrar):
                # Guarda os textos dos tópicos selecionados antes de limpar os checkbuttons
                selecionados_antes = set()
                for topic, var in var_checkbuttons:
                    if var.get():
                        selecionados_antes.add(topic['text'])

                # Limpa todos os checkbuttons
                for widget in frame_checkboxes.winfo_children():
                        widget.destroy()
                var_checkbuttons.clear()
                checkbox_widgets.clear()

                for topic in topicos_ordenados:  # Garante ordem inicial para consistência
                    if topic not in topicos_para_mostrar:
                        continue  # Só mostra os que passam no filtro
                    
                    texto = topic['text']
                    # Usa seleção salva ou False se ainda não tiver
                    selecionado = selecoes.get(texto, False)
                    var = tk.BooleanVar(value=selecionado)
                    def atualizar_selecao(t=texto, v=var):
                        selecoes[t] = v.get()
                    var.trace_add("write", lambda *args, t=texto, v=var: atualizar_selecao(t, v))
                    cb = tk.Checkbutton(frame_checkboxes, text=texto, variable=var, anchor='w', justify='left', wraplength=450)
                    cb.pack(anchor='w', pady=2)
                    var_checkbuttons.append((topic, var))
                    checkbox_widgets.append(cb)
                frame_checkboxes.update_idletasks()
                atualizar_scroll()
                canvas.yview_moveto(0)

            # Função de limpar pesquisa
            def limpar_pesquisa():
                campo_pesquisa.delete(0, tk.END)
                mostrar_topicos(topicos_ordenados)

            # Botão limpar (agora depois da função)
            botao_limpar = tk.Button(frame_pesquisa, text="Limpar", bg="#00579D", fg="white", command=limpar_pesquisa)
            botao_limpar.pack(side='right', padx=5)

            # Frame do canvas e scrollbar
            frame_canvas = tk.Frame(frame_principal)
            frame_canvas.pack(fill='both', expand=True)

            canvas = tk.Canvas(frame_canvas)
            scrollbar = tk.Scrollbar(frame_canvas, orient='vertical', command=canvas.yview)
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

            frame_checkboxes = tk.Frame(canvas)
            window_id = canvas.create_window((0, 0), window=frame_checkboxes, anchor='nw')

            
            def atualizar_scroll(event=None):
                root.update_idletasks()
                bbox = canvas.bbox("all")
                if not bbox:
                    return
            
                altura_conteudo = bbox[3] - bbox[1]
                altura_canvas = canvas.winfo_height()
            
                canvas.configure(scrollregion=bbox)
            
                if altura_conteudo <= altura_canvas:
                    scrollbar.pack_forget()
                    canvas.unbind_all("<MouseWheel>")
                else:
                    if not scrollbar.winfo_ismapped():
                        scrollbar.pack(side='right', fill='y')
                    canvas.bind_all("<MouseWheel>", on_mousewheel)

            def on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

            # Exibe todos os tópicos inicialmente
            mostrar_topicos(topicos_ordenados)

            def filtrar_topicos(event=None):
                texto_filtro = campo_pesquisa.get().lower()
                if texto_filtro.strip() == "":
                    mostrar_topicos(topicos_ordenados)
                else:
                    filtrados = [t for t in topicos_ordenados if texto_filtro in t['text'].lower()]
                    mostrar_topicos(filtrados)

            campo_pesquisa.bind("<KeyRelease>", filtrar_topicos)

            def obter_topicos_selecionados():
                topicos_selecionados = [
                    topicos_ordenados[i]['text']
                    for i in range(len(var_checkbuttons))
                    if var_checkbuttons[i][1].get()  # índice 1 = a variável BooleanVar
                ]
                if topicos_selecionados:
                    root.destroy()
                    processar_pdf(caminho_pdf, topicos_selecionados)
                else:
                    messagebox.showwarning("Aviso", "Nenhum tópico selecionado.")

            botao_criar_pdf = tk.Button(root, text="Criar PDF", bg="#00579D", fg="white", command=obter_topicos_selecionados)
            botao_criar_pdf.pack(pady=20)

            root.mainloop()

            if janela_foi_fechada:
                break

            continuar = messagebox.askyesno("Continuar", "Deseja criar outro PDF?")
            if not continuar:
                break

# Executar a função principal
caminho_pdf = selecionar_pdf()
criar_interface(caminho_pdf)
