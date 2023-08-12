def 로또()   :
    import random
    for i in range(1,6):
        a = random.sample(range(1,46),6)
        a.sort()
        print(a)

    # for i in range(2,10)    :
    #     for j in range(1,10) :
    #         hap = i * j
    #         print(hap,end=" ")
    #     print("")