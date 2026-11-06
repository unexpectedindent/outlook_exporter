"""Read-only discovery of Outlook accounts and data stores."""
from common import connect_outlook, list_stores, safe

def main():
    ns = connect_outlook()
    print('ACCOUNTS:')
    for i in range(1, ns.Accounts.Count+1):
        account = ns.Accounts.Item(i)
        store = safe(account, 'DeliveryStore')
        print(f'  {i}. {safe(account, "DisplayName")} | SMTP={safe(account, "SmtpAddress")} | delivery store={safe(store, "DisplayName")}')
    print('\nSTORES (copy exact names into additional_archive_store_names):')
    for store in list_stores(ns):
        print(f'  {safe(store,"DisplayName")} | FilePath={safe(store,"FilePath") or ""} | StoreID={(safe(store,"StoreID") or "")[:35]}...')

if __name__ == '__main__':
    main()
