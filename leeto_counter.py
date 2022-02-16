import json, requests, os
import pandas as pd
import numpy as np


def main():
    file_list = os.listdir()
    for file in file_list:
        if file.endswith('.xlsx'):
            filename = file
        else:
            exit

    df = pd.read_excel(filename)
    if 'Plate No' in df.columns:
        df.rename(columns={'Plate No': 'Bus No'}, inplace=True)
        print('[Rename] Plate No to Bus No')

    df['Stop Name'] = np.NaN
    df['Route'] = np.NaN

    for index in df.index:
        v = df.loc[index, 'GNSS']

        cords = v.split(',')
        cords_lat = cords[0]
        cords_lon = cords[1]
        if cords_lat.find('.') != -1:
            lat = cords_lat
            lon = cords_lon
        else:
            lat = cords_lat[:3] + '.' + cords_lat[3:]
            lon = cords_lon[:2] + '.' + cords_lon[2:]

        if is_lay_over(lat, lon):
            df.loc[index, 'Stop Name'] = 'LayOver'
            df.loc[index, 'Route'] = 'LayOver'
            print('[Updating] Layover. Row: ' + str(index))
            continue
        elif is_fuel_wise(lat, lon):
            df.loc[index, 'Stop Name'] = 'Fuel Wise'
            print('[Updating] Fuel wise. Row: ' + str(index))

        url = 'http://46.101.72.176:8080/otp/routers/default/index/stops?lat=' + lat + '&lon=' + lon + '&radius=300'

        try:
            response = requests.get(url)
        except requests.exceptions.ConnectionError:
            print('[Error] Lost connection with the server. Row: ' + str(index))
            continue

        if response.status_code == 200:
            data_res = json.loads(response.text)
            if len(data_res) > 1:
                dist_list = []
                for i in range(len(data_res)):
                    dist_list.append(data_res[i]['dist'])
                min_dist = min(dist_list)
                min_index = dist_list.index(min_dist)
                stop_name = data_res[min_index]['name']
                stop_id = data_res[min_index]['id']
            elif len(data_res) == 1:
                stop_name = data_res[0]['name']
                stop_id = data_res[0]['id']
            else:
                continue
        else:
            continue

        church = ['410', '316', '409', '509']
        if stop_name in church:
            stop_name = 'Church Street'
            route = ''
        else:
            str_url = 'http://46.101.72.176:8080/otp/routers/default/index/stops/' + stop_id + '/routes'
            res = requests.get(str_url)
            if res.status_code == 200:
                j = json.loads(res.text)
                if len(j) >= 1:
                    route = j[0]['shortName'] + ' - ' + j[0]['longName']
                else:
                    continue

            else:
                continue

        df.loc[index, 'Stop Name'] = stop_name
        df.loc[index, 'Route'] = route
        print('[Updating] Row: ' + str(index) + ' Stop Name : ' + stop_name + ' Route: ' + route)
    filename_l = filename.split('.')
    new_file_name = filename_l[0] + ' updated.' + filename_l[1]
    df = df[['Bus No', 'IN', 'Number of people', 'Alarm Time', 'GNSS', 'Stop Name', 'Route']]
    print('[Adding] Dates')
    df = add_date(df)
    df['Week'] = df['DOW'].apply(add_week)
    print('[Saving] df to excel')
    df.to_excel(new_file_name, index=False)
    print('Done.')


def add_date(df):
    df['Date'] = pd.to_datetime(df['Alarm Time'])
    df['Year'] = df.Date.dt.year.astype(str)
    df['Month'] = df.Date.dt.month_name()
    df['Day'] = df.Date.dt.day.astype(str)
    df['DOW'] = df.Date.dt.day_name()
    df['MY'] = df.Month + ' ' + df.Year

    return df


def add_week(name):
    weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    weekends = ['Sunday', 'Saturday']
    if name in weekends:
        return 'Weekend'
    elif name in weekdays:
        return 'Weekday'


def is_fuel_wise(arg_lat, arg_lon):
    arg_lat = float(arg_lat)
    arg_lon = float(arg_lon)

    min_lat = -23.89453
    max_lat = -23.89581
    min_lon = 29.44261
    max_lon = 29.44393

    if max_lat < arg_lat < min_lat:
        if min_lon < arg_lon < max_lon:
            return True
        else:
            return False

    else:
        return False


def is_lay_over(arg_lat, arg_lon):
    arg_lat = float(arg_lat)
    arg_lon = float(arg_lon)

    min_lat = -23.89813
    max_lat = -23.89937
    min_lon = 29.44261
    max_lon = 29.44507

    if max_lat < arg_lat < min_lat:
        if min_lon < arg_lon < max_lon:
            return True
        else:
            return False

    else:
        return False


if __name__ == '__main__':
    main()
