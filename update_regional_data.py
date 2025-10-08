import pandas as pd
from pykakasi import kakasi
import re

def parse_full_name(full_name_str):
    """
    カンマ区切りの英語の正式名をパースして、地名と都道府県名を返す。
    例: "Kesennuma, Miyagi, Japan" -> ("Kesennuma", "Miyagi")
    """
    parts = [p.strip() for p in full_name_str.split(',')]
    if len(parts) >= 2:
        return parts[0], parts[1]
    return parts[0], None

def finalize_name(city_name_from_list, region_type):
    """
    市区町村リストから取得した名前に、正しい接尾語（市・町・区・村）がついているか確認し、
    なければ地域種別に応じて付与する。郡が含まれている場合は、郡を除いた部分を返す。
    """
    # 郡が含まれている場合、それを取り除く
    if '郡' in city_name_from_list:
        match = re.search(r'郡(.+)', city_name_from_list)
        if match:
            city_name_from_list = match.group(1)

    # 接尾語を決定
    suffix_map = {
        'City': '市',
        'Town': '町',
        'Ward': '区',
        'Village': '村',
        'Prefecture': '県'
    }
    suffix = suffix_map.get(region_type, '')

    # 既に正しい接尾語がついているかチェック
    if not city_name_from_list.endswith(suffix):
        # ついていない場合、もし他の接尾語がついていたら削除
        for s in suffix_map.values():
            if city_name_from_list.endswith(s):
                city_name_from_list = city_name_from_list[:-len(s)]
                break
        # 正しい接尾語を付与
        return city_name_from_list + suffix

    return city_name_from_list

def update_csv_final_attempt(template_path, city_list_path, output_path):
    """
    ユーザーの最終要件に基づき、CSVを更新する。
    - 名前: ◯◯市, ◯◯町, ◯◯区 のみ
    - 正式名: 都道府県から始まるフルネーム
    """
    try:
        # --- 準備フェーズ ---
        k = kakasi()
        k.setMode("H", "a")
        k.setMode("K", "a")
        k.setMode("J", "a")
        conv = k.getConverter()

        template_df = pd.read_csv(template_path)
        city_list_df = pd.read_csv(city_list_path, header=None, names=['コード', '市区町村名', '読みガナ', '都道府県コード', '都道府県名', '都道府県読みガナ'])

        city_data = []
        for _, row in city_list_df.iterrows():
            city_name_no_gun = re.sub(r'.+郡', '', row['市区町村名'])
            city_data.append({
                '市区町村名_フル': row['市区町村名'],
                '市区町村名_郡なし': city_name_no_gun,
                '都道府県名': row['都道府県名'],
                'romaji_city': conv.do(city_name_no_gun).upper().replace(' ', ''),
                'romaji_pref': conv.do(row['都道府県名']).upper().replace(' ', '')
            })

        city_data_df = pd.DataFrame(city_data)

        # --- 更新フェーズ ---
        for index, row in template_df.iterrows():
            if row['地域種別'] not in ['City', 'Town', 'Ward', 'Village', 'District', 'City region']:
                continue

            eng_name, eng_pref = parse_full_name(str(row['正式名']))
            search_name_romaji = eng_name.upper().replace(' ', '').replace('-', '')
            search_pref_romaji = eng_pref.upper().replace(' ', '') if eng_pref else None

            candidates = city_data_df[city_data_df['romaji_city'].str.contains(search_name_romaji, na=False)]

            best_match = None
            if not candidates.empty:
                if len(candidates) > 1 and search_pref_romaji:
                    # 都道府県名で絞り込み
                    pref_candidates = candidates[candidates['romaji_pref'].str.contains(search_pref_romaji, na=False)]
                    if not pref_candidates.empty:
                        best_match = pref_candidates.iloc[0]

                if best_match is None:
                    # 絞り込めなかったか、都道府県情報がなかったので、最初の候補を選ぶ
                    best_match = candidates.iloc[0]

            if best_match is not None:
                # 名前列の生成
                final_name = finalize_name(best_match['市区町村名_郡なし'], row['地域種別'])

                # 正式名列の生成
                formal_name = best_match['都道府県名'] + best_match['市区町村名_フル']

                # 更新
                template_df.at[index, '名前'] = final_name
                template_df.at[index, '正式名'] = formal_name
                print(f"'{row['名前']}' -> 名前: '{final_name}', 正式名: '{formal_name}'")

        template_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\n更新完了！新しいファイル'{output_path}'が作成されました。")

    except FileNotFoundError as e:
        print(f"エラー: {e}")
    except Exception as e:
        print(f"予期せぬエラー: {e}")

if __name__ == '__main__':
    template_file = '地域リスト/【テンプレート_v2.2】◯◯様_Google広告データ取得用 - 地域マスターリスト.csv'
    city_list_file = '地域リスト/市区町村リスト.csv'
    output_file = '地域リスト/【更新版】地域マスターリスト.csv'
    update_csv_final_attempt(template_file, city_list_file, output_file)
