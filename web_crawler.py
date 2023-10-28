import time
import threading

#設定平行運行
def subroutine_1():
    import dell_NB
    time.sleep(3)
    import dell_DT
    time.sleep(3)
    import Lenovo_DT
    
def subroutine_2():
    time.sleep(2)
    import Lenovo_NB

def subroutine_3():
    time.sleep(3)
    import HP_NB
    time.sleep(3)
    import HP_DT
    
def subroutine_4():
    time.sleep(4)
    import Lenovo_docking
    time.sleep(3)
    import HP_docking    
    time.sleep(3)
    import dell_dock
    
def subroutine_5():
    import reorganize_data
    
#分配運行核心
Subroutine_1 = threading.Thread(target=subroutine_1)
Subroutine_2 = threading.Thread(target=subroutine_2)
Subroutine_3 = threading.Thread(target=subroutine_3)
Subroutine_4 = threading.Thread(target=subroutine_4)
Subroutine_5 = threading.Thread(target=subroutine_5)

#開始執行
Subroutine_1.start()
Subroutine_2.start()
Subroutine_3.start()
Subroutine_4.start()

#等待執行完畢
Subroutine_1.join()   # 加入等待 aa() 完成的方法
Subroutine_2.join()   # 加入等待 bb() 完成的方法
Subroutine_3.join()   # 加入等待 cc() 完成的方法
Subroutine_4.join()   # 加入等待 dd() 完成的方法

#執行最後程序
Subroutine_5.start()  # 當前面都執行完，才會開始執行


