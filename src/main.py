from google.cloud import ndb
from flask import Flask, request
from os import environ
import logging


class Room(ndb.Model):
    name = ndb.StringProperty()
    capacity = ndb.IntegerProperty()
    schedule = ndb.StringProperty()
    takers = ndb.StringProperty()


client = ndb.Client()


# Middleware to obtain new client context for each request. This code borrowed from Google
# at https://cloud.google.com/appengine/docs/standard/python3/migrating-to-cloud-ndb
def ndb_wsgi_middleware(wsgi_app):
    def middleware(environ, start_response):
        with client.context():
            return wsgi_app(environ, start_response)

    return middleware


app = Flask(__name__)
# Wrap app in middleware
app.wsgi_app = ndb_wsgi_middleware(app.wsgi_app)


@app.route('/list')
def list_rooms():
    rooms = [
        {
            'name': r.name,
            'capacity': r.capacity,
            'schedule': r.schedule,
            'takers': r.takers,
            'key': r.key.urlsafe().decode('utf-8')
        } for r in Room.query().order(Room.name)]

    resp = ''
    for r in rooms:
        resp += '|'.join([r['key'], r['name'], str(r['capacity']), r['schedule'], r['takers']]) + ';'

    return resp


@app.route('/save', methods=['POST'])
def save_room():
    params = request.get_json(force=True)
    name = params['name']
    capacity = params['capacity']
    schedule = params['schedule']
    takers = params['takers']

    room = Room(name=name, capacity=capacity, schedule=schedule, takers=takers)
    key = room.put()
    return 'roomkey=' + key.urlsafe().decode('utf-8')


@app.route('/update', methods=['POST'])
@ndb.transactional(retries=0)
def reserve_room():
    params = request.get_json(force=True)
    roomkey = params['roomkey']
    schedule = params['schedule']
    taker = params['taker']

    try:
        room = ndb.Key(urlsafe=roomkey).get()

        allok = True
        for idx, slot in enumerate(schedule):
            if slot == '1':
                if room.schedule[idx] == '0':
                    x = list(room.schedule)
                    x[idx] = '1'
                    room.schedule = ''.join(x)

                    x = room.takers.split(':')
                    logging.debug('PJS: idx={0}, x length={1}'.format(idx, len(x)))
                    x[idx] = taker
                    room.takers = ':'.join(x)
                else:
                    allok = False
                    break

        if allok:
            room.put()
            resp = 'roomkey=' + roomkey
        else:
            resp = 'error=Room not free at requested times'
    except Exception as ex:
        resp = f'error=Could not reserve room for requested times; {ex}'

    return resp


@app.route('/purge')
def purge_all():
    keys = [r.key for r in Room.query()]

    try:
        ndb.delete_multi(keys)
        resp = 'All rooms deleted'
    except Exception as ex:
        resp = f'Error purging rooms: {ex}'

    return resp
