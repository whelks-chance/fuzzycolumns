import pprint

from fuzzy_columns import FuzzyColumns

if __name__ == '__main__':

    fc = FuzzyColumns()
    results = fc.compare_spreadsheets(
        '/home/ianh/Downloads/metadata_2.xlsx',
        '/home/ianh/dublin_core_headers.xlsx'
    )
    # print(pprint.pformat(results))
    print('\n\n')
    print(pprint.pformat(results.best_matches()))
    print('\n\n')
    print(results.best_matches())
    print('\n\n')
    print(results.best_matches()[''])
    print('\n\n')
    print(results.best_matches()['matches']['matches'])
    # print(results.print_distribution())