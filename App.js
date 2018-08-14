// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};
		var prevHeader;

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
					prevHeader = null;
          $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
						const $notifBodyElem = $('#notification-message-body');
						const prevText = prevHeader === header ? `${$notifBodyElem.html()}<br />` : ''
						prevHeader = header;

            $('#notification-message-header').text(header);
            $notifBodyElem.html(`${prevText}${text}`);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();
