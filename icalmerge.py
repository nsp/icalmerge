import click
import csv
from datetime import datetime
import icalendar
from openpyxl import load_workbook
import os
from pytz import timezone


SOURCE_SHEET = u'MailMergeSource'
SCHED_SHEET = u'CampSchedule'
INSTR_SHEET = u'Instructors'
OUT_SHEET = u'MailMergeCalendars'

@click.command()
@click.argument('workbook', type=click.File('rb'))
@click.argument('outdir', type=click.Path(exists=False, file_okay=False,
                                          writable=True))
def cli(workbook, outdir):
    wb = load_workbook(workbook)
    assert wb.get_sheet_names() == [SOURCE_SHEET, SCHED_SHEET, INSTR_SHEET, OUT_SHEET]
    works = [row for row in wb.get_sheet_by_name(SOURCE_SHEET).rows[1:]]
    instr_locs = {instr.value: loc.value for instr, loc
                  in wb.get_sheet_by_name(INSTR_SHEET).rows[1:]}
    sched = [row for row in wb.get_sheet_by_name(SCHED_SHEET).rows[1:]]
    if not os.path.isdir(outdir):
        os.mkdir(outdir)
    now = datetime.utcnow()
    for student, a, b, c, d, email, phone in works:
        instrs = [a.value, b.value, c.value, d.value]
        if student.value is None:
            break
        click.echo('Creating schedule for "{}"'.format(student.value))
        cal = icalendar.Calendar(prodid='-//nsp//icalmerge 0.1',
                                 version='2.0',
                                 calscale='GREGORIAN')
        cal.add('x-wr-calname', 'Camp Improv Utopia 2016 Schedule for {}'.format(student.value))
        cal.add('x-wr-caldesc', 'All major events and workshops.')
        wkshp_idx = 0
        for what, where, start, end, tz in sched:
            ev = icalendar.Event(uid='{}@improvutopia.com'.format(what.value),
                                 dtstamp=now,
                                 created=now)
            ev.add('last-modified', now)
            ev.add('dtstart', start.value.replace(tzinfo=timezone(tz.value)))
            ev.add('dtend', end.value.replace(tzinfo=timezone(tz.value)))
            if where.value is None and instrs[wkshp_idx] is not None:
                ev.add('summary', '{} with {}'.format(what.value, instrs[wkshp_idx]))
                ev.add('description', instrs[wkshp_idx])
                ev.add('location', instr_locs[instrs[wkshp_idx]])
                wkshp_idx += 1
            else:
                ev.add('summary', what.value)
                ev.add('location', where.value)
            cal.add_component(ev)
        with open(os.path.join(outdir, '{}.ics'.format(student.value)), 'wb') as f:
            f.write(cal.to_ical())


if __name__ == '__main__':
    cli()
