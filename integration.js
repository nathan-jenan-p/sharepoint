let async = require('async');
let config = require('./config/config');
let request = require('request');
let util = require('util');

let Logger;
let requestOptions = {};

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

function getRequestOptions() {
    return JSON.parse(JSON.stringify(requestOptions));
}

function getAuthToken(options, callback) {
    request({
        url: `https://${options.authHost}/${options.tenantId}/tokens/OAuth/2`,
        formData: {
            grant_type: 'client_credentials',
            client_id: `${options.clientId}@${options.tenantId}`,
            client_secret: options.clientSecret,
            resource: `00000003-0000-0ff1-ce00-000000000000/${options.host}@${options.tenantId}`,
        },
        json: true,
        method: 'POST'
    }, (err, resp, body) => {
        if (err) {
            callback(err);
            return;
        }

        if (resp.statusCode != 200) {
            callback({ err: new Error('status code was not 200'), body: body });
            return;
        }

        callback(null, body.access_token);
    });
}

function querySharepoint(entity, token, options, callback) {
    let requestOptions = getRequestOptions();
    requestOptions.qs = {
        querytext: `'${entity.value}'`
    };
    requestOptions.url = `https://${options.host}/_api/search/query`;
    requestOptions.headers = {
        Authorization: 'Bearer ' + token
    };

    request(requestOptions, (err, resp, body) => {
        if (err || resp.statusCode != 200) {
            callback(err || new Error('status code was ' + resp.statusCode));
            return;
        }

        callback(null, body);
    });
}

function doLookup(entities, options, callback) {
    Logger.trace('starting lookup');

    Logger.trace('options are', options);

    let results = [];

    getAuthToken(options, (err, token) => {
        if (err) {
            Logger.error('get token errored', err);
            callback({ err: err });
            return;
        }

        // We have to do 1 request per query because we can only AND the query 
        // params not OR them
        async.each(entities, (entity, done) => {
            querySharepoint(entity, token, options, (err, body) => {
                if (err) {
                    done(err);
                    return;
                }

                if (body.PrimaryQueryResult.RelevantResults.RowCount < 1) {
                    results.push({
                        entity: entity,
                        data: null
                    });
                    done();
                    return;
                }

                results.push({
                    entity: entity,
                    data: {
                        summary: [],
                        details: formatSearchResults(body)
                    }
                });
                done();
            });
        }, err => {
            if (err) {
                Logger.error('lookup errored', err);

                // errors can sometime have circular structure and this breaks polarity
                callback({ err: util.inspect(err) });
                return;
            }

            Logger.trace('sending results to client', results);

            callback(null, results);
        });
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

    validateStringOption(errors, options, 'host', 'You must provide a Host option.');
    validateStringOption(errors, options, 'clientId', 'You must provide a Client ID option.');
    validateStringOption(errors, options, 'clientSecret', 'You must provide a Client Secret option.');
    validateStringOption(errors, options, 'tenantId', 'You must provide a Tenant ID option.');

    callback(null, errors);
}

module.exports = {
    doLookup: doLookup,
    startup: startup,
    validateOptions: validateOptions
};
