let async = require('async');
let config = require('./config/config');
let request = require('request');

let Logger;
let requestWithDefaults;
let requestOptions = {};

let searchData = require('./search');

function handleRequestError(request) {
    return (options, expectedStatusCode, callback) => {
        return request(options, (err, resp, body) => {
            if (err || resp.statusCode !== expectedStatusCode) {
                Logger.error(`error during http request to ${options.url}`, { error: err, status: resp ? resp.statusCode : 'unknown' });
                callback({ error: err, statusCode: resp ? resp.statusCode : 'unknown' });
            } else {
                callback(null, body);
            }
        });
    };
}

function formatSearchResults(searchResults) {
    let fakeData = [];

    searchResults.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(row => {
        let obj = {};
        row.Cells.forEach(cell => {
            obj[cell.Key] = cell.Value;
        });

        fakeData.push(obj);
    });

    return fakeData
}

function querySharepoint(entity, options, callback) {
    if (options.fakeData) {
        callback(null, {
            entity: entity,
            data: {
                summary: [],
                details: formatSearchResults(searchData)
            }
        });
        return;
    }

    requestWithDefaults({
        url: `https://${options.host}/_api/search/query`,
        method: 'GET',
        qs: {
            querytext: `'${entity.value}'`
        },
        // TODO figure out authentication
        headers: {
            'Cookie': require('./creds').cookie,
        }
    }, 200, (err, resp) => {
        if (err) {
            callback(err);
            return;
        }

        if (resp.PrimaryQueryResult.RelevantResults.RowCount < 1) {
            callback(null, {
                entity: entity,
                data: null,
            });
            return;
        }

        callback(null, {
            entity: entity,
            data: {
                summary: [],
                details: formatSearchResults(resp)
            }
        });
    });
}

function doLookup(entities, options, callback) {
    // We have to do 1 request per query because we can only AND the query 
    // params not OR them
    Logger.trace('starting lookup')

    let results = [];

    async.each(entities, (entity, done) => {
        querySharepoint(entity, options, (err, resp) => {
            if (err) {
                done(err);
                return;
            }

            results.push(resp);
            done();
        });
    }, err => {
        if (err) {
            callback(err);
            return;
        }

        Logger.trace('sending results to client', results);

        callback(null, results);
    });
}

function startup(logger) {
    Logger = logger;

    if (typeof config.request.cert === 'string' && config.request.cert.length > 0) {
        requestOptions.cert = fs.readFileSync(config.request.cert);
    }

    if (typeof config.request.key === 'string' && config.request.key.length > 0) {
        requestOptions.key = fs.readFileSync(config.request.key);
    }

    if (typeof config.request.passphrase === 'string' && config.request.passphrase.length > 0) {
        requestOptions.passphrase = config.request.passphrase;
    }

    if (typeof config.request.ca === 'string' && config.request.ca.length > 0) {
        requestOptions.ca = fs.readFileSync(config.request.ca);
    }

    if (typeof config.request.proxy === 'string' && config.request.proxy.length > 0) {
        requestOptions.proxy = config.request.proxy;
    }

    if (typeof config.request.rejectUnauthorized === 'boolean') {
        requestOptions.rejectUnauthorized = config.request.rejectUnauthorized;
    }

    requestOptions.json = true;

    requestWithDefaults = handleRequestError(request.defaults(requestOptions));
}

function validateStringOption(errors, options, optionName, errMessage) {
    if (typeof options[optionName].value !== 'string' ||
        (typeof options[optionName].value === 'string' && options[optionName].value.length === 0)) {
        errors.push({
            key: optionName,
            message: errMessage
        });
    }
}

function validateOptions(options, callback) {
    let errors = [];

    // Example of how to validate a string option
    validateStringOption(errors, options, 'host', 'You must provide a host option.');

    callback(null, errors);
}

module.exports = {
    doLookup: doLookup,
    startup: startup,
    validateOptions: validateOptions
};
