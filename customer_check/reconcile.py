from collections import defaultdict

from .models import Entry, Result, Table


def normalize(value: str) -> str:
    return value.strip().casefold()


def compare(
    registry: Table, payments: Table, registry_code: int,
    status_column: int, payment_code: int, active_values: set[str],
    name_column: int | None = None,
) -> Result:
    for table, columns in [(registry, [registry_code, status_column, name_column]), (payments, [payment_code])]:
        if any(column is not None and not 0 <= column < len(table.headers) for column in columns):
            raise ValueError("Επίλεξε έγκυρες στήλες για τα δύο αρχεία.")
    if registry_code == status_column:
        raise ValueError("Ο κωδικός πελάτη και η κατάσταση συμβολαίου πρέπει να είναι διαφορετικές στήλες.")
    active = {normalize(value) for value in active_values if value.strip()}
    if not active:
        raise ValueError("Επίλεξε τουλάχιστον μία τιμή που σημαίνει ενεργό συμβόλαιο.")
    available = {normalize(row.values[status_column]) for row in registry.rows}
    if not active.issubset(available):
        raise ValueError("Κάποιες ενεργές τιμές δεν υπάρχουν στο επιλεγμένο μητρώο.")
    result = Result(registry_rows=len(registry.rows), payment_rows=len(payments.rows))
    grouped = defaultdict(list)
    for row in registry.rows:
        code = row.values[registry_code]
        name = row.values[name_column] if name_column is not None else ""
        status = row.values[status_column]
        entry = Entry(code, name, row.number)
        if not normalize(code):
            result.issues.append(Entry(code, name, row.number, "Μητρώο: κενός κωδικός πελάτη"))
            continue
        grouped[normalize(code)].append((entry, normalize(status)))
    payment_codes = set()
    payment_entries = {}
    for row in payments.rows:
        code = row.values[payment_code]
        key = normalize(code)
        if not key:
            result.issues.append(Entry("", "", row.number, "Πληρωμές: κενός κωδικός πελάτη"))
            continue
        if key in payment_codes:
            result.duplicate_payments += 1
        payment_codes.add(key)
        payment_entries.setdefault(key, Entry(code, "", row.number, "Πληρωμές: ο κωδικός δεν υπάρχει στο μητρώο"))
    for key, records in grouped.items():
        if len(records) > 1:
            for entry, status in records:
                result.issues.append(Entry(entry.code, entry.name, entry.source_row, "Μητρώο: διπλός κωδικός — απαιτείται έλεγχος συμβολαίων"))
            continue
        entry, status = records[0]
        if not status:
            result.issues.append(Entry(entry.code, entry.name, entry.source_row, "Μητρώο: κενή κατάσταση συμβολαίου"))
        elif status in active:
            (result.found if key in payment_codes else result.missing).append(entry)
        else:
            result.inactive_rows += 1
    result.unknown = [payment_entries[key] for key in sorted(payment_codes - grouped.keys())]
    return result

