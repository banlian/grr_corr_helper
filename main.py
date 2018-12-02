import grrhelper


if __name__ == '__main__':

    print('''
    --------------------------------------------------
                grr corr helper
     --------------------------------------------------
    1. input calc type:
        grr
        corr            
    
    2. input product type:
        p1
    
    3. input station type

    
    4. input raw data file
    --------------------------------------------------
    ''')


    while(True):
        print('input calc type:')
        calc = input()

        print('\ninput product type:')
        pro = input()

        print('\ninput station type:')
        s = input()

        print('\ninput data file:')
        data = input()


        product = pro
        raw_data_file = data
        station = s
        if calc == 'grr':
            print("run grr %s %s %s start\n" %(product, station, raw_data_file))
            grr = grrhelper.GrrHelper()
            grr.create(product, station)
            grr.run_with_data(raw_data_file)
            grr.display()
            print("run grr %s %s %s finish\n" % (product, station, raw_data_file))
        elif calc == 'corr':
            pass
        else:
            print('not support calc %s' %(calc))


        #print continue?
        print('\ncontinue? Y : N')
        choice = input()
        if choice == 'Y':
            continue
        else:
            break


    print('exit.............')
    input()