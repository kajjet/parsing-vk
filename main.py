import vk
import time
import xlwt
import os

V = '5.154'

token = "ввод от пользователя"
api = vk.API(access_token=token, v=V)

members = api.groups.getMembers(group_id='tortoreto')
count = members['count']
offset = 1000
members = members['items']

while offset < count:
    members.extend(api.groups.getMembers(group_id='tortoreto', count=1000, offset=offset)['items'])
    offset += 1000
    if offset % 10000 == 0:
        time.sleep(1)

if os.path.isfile('data.xls'):
    os.remove('data.xls')

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

i = 1

sheet1.write(0, 0, 'Имя')
sheet1.write(0, 1, 'Фамилия')
sheet1.write(0, 2, 'Возраст')
sheet1.write(0, 3, 'Пол')
sheet1.write(0, 4, 'Интересы')

request_count = 0

categories_count = {}

for user_id in members:

    user = api.users.get(user_id=user_id, fields=['sex', 'bdate', 'deactivated'])[0]
    request_count += 1
    if request_count % 5 == 0:
        time.sleep(2)
    if user['is_closed'] or user.get('deactivated'):
        continue

    groups = api.groups.get(user_id=user_id)
    count = groups['count']
    offset = 1000
    groups = groups['items']
    while offset < count:
        groups.extend(api.groups.get(user_id=user_id, count=1000, offset=offset)['items'])
        offset += 1000
        if offset % 10000 == 0:
            time.sleep(1.5)

    categories = {}

    for group_indexes in range(0, len(groups), 500):
        selected_groups = api.groups.getById(group_ids=groups[group_indexes:group_indexes + 500], fields=['activity'])
        for group in selected_groups['groups']:
            if categories.get(group.get('activity')):
                categories[group.get('activity')] += 1

            else:
                categories[group.get('activity')] = 1
                if group.get('activity') is None:
                    del categories[None]

                elif 'заблокирован' in group.get('activity'):
                    del categories[group.get('activity')]

                else:
                    for number in '0123456789':
                        if number in group.get('activity'):
                            del categories[group.get('activity')]
                        break

    if categories.get('Закрытое сообщество'):
        del categories['Закрытое сообщество']

    info = sorted(categories.items(), key=lambda x: x[1], reverse=True)

    if user['sex'] == 0:
        user['sex'] = 'пол не указан'

    elif user['sex'] == 1:
        user['sex'] = 'женский'

    else:
        user['sex'] = 'мужской'

    if not user.get('bdate') or not len(user.get('bdate').split('.')) == 3:
        user['bdate'] = 'не указан'
    else:
        user['bdate'] = user['bdate'].split('.')[2]
    if len(info) == 0:
        data = (
            user.get('first_name'), user.get('last_name'), user.get('bdate'), user.get('sex'), 'Отсутствуют')
    elif len(info) == 1:
        data = (
            user.get('first_name'), user.get('last_name'), user.get('bdate'), user.get('sex'), info[0][0])
        if categories_count.get(info[0][0]):
            categories_count[info[0][0]] += 1

        else:
            categories_count[info[0][0]] = 1
    else:
        data = (
            user.get('first_name'), user.get('last_name'), user.get('bdate'), user.get('sex'),
            f'{info[0][0]}, {info[1][0][0].lower() + info[1][0][1:]}')

        if categories_count.get(info[0][0]):
            categories_count[info[0][0]] += 1

        else:
            categories_count[info[0][0]] = 1

        if categories_count.get(info[1][0]):
            categories_count[info[1][0]] += 1

        else:
            categories_count[info[1][0]] = 1

    for index in range(len(data)):
        sheet1.write(i, index, data[index])


    i += 1
    request_count += 1
    if request_count % 5 == 0:
        time.sleep(2)

categories_count = sorted(categories_count.items(), key=lambda x: x[1], reverse=True)

print(categories_count)

sheet1.write(0, 8, "Интерес")
sheet1.write(0, 9, "Процент заинтересованных")
for i in range(len(categories_count)):
    sheet1.write(i + 1, 8, categories_count[i][0])
    print(round(categories_count[i][1] / len(categories_count), 2))
    sheet1.write(i + 1, 9, f'{int(round(categories_count[i][1] / len(categories_count) * 100, 2))}%')
print(1)
book.save("data.xls")
