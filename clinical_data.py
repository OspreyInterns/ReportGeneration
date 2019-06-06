
import sqlite3 as sqlite

# Reads from the injection table to sum up the injections


def straight_to_patient(case_number: int, file_name):

    _con = sqlite.connect(file_name)

    with _con:
        contrast_inj = 0.
        mismatch = False
        _cur = _con.cursor()
        _cur.execute('SELECT * FROM CMSWInjections')

        _cols = _cur.fetchall()

        for _col in _cols:
            # _col[18](%) matches Alex's data, _col[17](volume) goes by volume diverted
            if _col[1] == case_number and _col[5] == 1 and _col[18] == 0:
                contrast_inj += _col[20]
                if _col[17] != 0:
                    print('Case', _col[1], 'contains a mismatch between % and volume diverted')
                    mismatch = True
                # _col[12] = total injection _col[16] = diverted volume _col[19] = total volume to patient
                # _col[30] = pressure _col[32] = pause _col[29] = flow rate to patient
                if round(_col[12], 4) != round(_col[16] + _col[19], 4) and _col[30] == 0 and _col[32] == 0:
                    if _col[29] != 0:
                        print('Injection', _col[0], 'suspicious', _col[12], '!=', _col[16] + _col[19])

        return [contrast_inj, mismatch]
