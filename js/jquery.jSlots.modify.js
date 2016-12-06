/*
 * jQuery jSlots Plugin
 * http://matthewlein.com/jslot/
 * Copyright (c) 2011 Matthew Lein
 * Version: 1.0.2 (7/26/2012)
 * Dual licensed under the MIT and GPL licenses
 * Requires: jQuery v1.4.1 or later
 */

(function ($) {

    $.jSlots = function (el, options) {

        var base = this;

        base.$el = $(el);
        base.el = el;

        base.$el.data("jSlots", base);

        base.init = function () {

            base.options = $.extend({}, $.jSlots.defaultOptions, options);

            base.setup();
            //base.bindEvents();

        };


        // --------------------------------------------------------------------- //
        // DEFAULT OPTIONS
        // --------------------------------------------------------------------- //

        $.jSlots.defaultOptions = {
            number: 3,          // Number: number of slots
            //winnerNumber: 1,    // Number or Array: list item number(s) upon which to trigger a win, 1-based index, NOT ZERO-BASED
            //spinner: '',        // CSS Selector: element to bind the start event to
            //spinEvent: 'click', // String: event to start slots on this event
            onStart: $.noop,    // Function: runs on spin start,
            onEnd: $.noop,      // Function: run on spin end. It is passed (finalNumbers:Array). finalNumbers gives the index of the li each slot stopped on in order.
            //onWin: $.noop,      // Function: run on winning number. It is passed (winCount:Number, winners:Array)
            //easing: 'easeOutSine',    // String: easing type for final spin
            //time: [2000, 4000, 7000],         // Number: total time of spin animation
            loops: [2, 4, 7],            // Number: times it will spin during the animation            
            min: 1,
            max: 127,
        };

        // --------------------------------------------------------------------- //
        // HELPERS
        // --------------------------------------------------------------------- //

        base.randomRange = function (min, max, exclude) {
            var num = Math.floor(Math.random() * (max - min + 1)) + min;
            return exclude.indexOf(num) === -1 ? num : this.randomRange(min, max, exclude);
        }

        // --------------------------------------------------------------------- //
        // VARS
        // --------------------------------------------------------------------- //

        base.isSpinning = false;
        base.spinSpeed = 0;
        //base.winCount = 0;
        base.doneCount = 0;

        base.$liHeight = 0;
        base.$liWidth = 0;

        //base.winners = [];
        base.allSlots = [];
        base.exclude = [];
        base.endNum = [];

        // --------------------------------------------------------------------- //
        // FUNCTIONS
        // --------------------------------------------------------------------- //


        base.setup = function () {

            // set sizes

            var $list = base.$el;
            var $li = $list.find('li').first();

            base.$liHeight = $li.outerHeight();
            base.$liWidth = $li.outerWidth();

            console.log("height : " + base.$liHeight + " width : " + base.$liWidth);

            base.liCount = base.$el.children().length;

            base.listHeight = base.$liHeight * base.liCount;

            //base.increment = (base.options.time / base.options.loops) / base.options.loops;

            $list.css('position', 'relative');

            $li.clone().appendTo($list);

            base.$wrapper = $list.wrap('<div class="jSlots-wrapper"></div>').parent();

            // remove original, so it can be recreated as a Slot
            base.$el.remove();

            // clone lists
            for (var i = 0; i <= base.options.number - 1; i++) {
                base.allSlots.push(new base.Slot(i));
            }

        };

        //base.bindEvents = function () {
        //    $(base.options.spinner).bind(base.options.spinEvent, function (event) {
        //        if (!base.isSpinning) {
        //            base.playSlots();
        //        }
        //    });
        //};

        // Slot contstructor
        base.Slot = function (i) {

            this.index = i;
            this.spinSpeed = 0;
            this.increment = 0;//(base.options.time[i] / base.options.loops[i]) / base.options.loops[i];
            this.el = base.$el.clone().appendTo(base.$wrapper)[0];
            this.$el = $(this.el);
            this.loopCount = 0;
            //this.number = 0;

        };


        base.Slot.prototype = {

            // do one rotation
            spinEm: function () {

                var that = this;

                //console.log("animationSpeed : " + that.spinSpeed);

                that.$el
                    .css('top', -base.listHeight)
                    .animate({ 'top': '0px' }, that.spinSpeed, 'linear', function () {
                        that.lowerSpeed();
                    });

            },

            lowerSpeed: function () {

                //console.log("animationSpeed : " + that.spinSpeed);

                this.spinSpeed += this.increment;
                this.loopCount++;

                if (this.loopCount < base.options.loops[this.index]) {

                    this.spinEm();

                } else {

                    this.finish();

                }
            },

            // final rotation
            finish: function () {

                var that = this;
                //

                var finalPos = -((base.$liHeight * (base.endNum[that.index] + 1)) - base.$liHeight);
                var finalSpeed = ((this.spinSpeed * 0.5) * (base.liCount + this.index * 2)) / (base.endNum[that.index] + 1);

                finalSpeed = finalSpeed * (this.index * 2 + 5) / 10;

                //console.log('index : ' + this.index + 'finalSpeed : ' + finalSpeed)

                that.$el
                    .css('top', -base.listHeight)
                    .animate({ 'top': finalPos }, finalSpeed, 'easeOutSine', function () {
                        base.endSlot();
                    });

            },

            noAnimation: function () {
                var that = this;

                var finalPos = -((base.$liHeight * (base.endNum[that.index] + 1)) - base.$liHeight);

                that.$el.css('top', finalPos);

                base.endSlot();
                //that.$el
                //    .css('top', -base.listHeight)
                //    .animate({ 'top': finalPos }, 100, 'linear', function () {
                //        base.endSlot();
                //    });
            }

        };

        base.endSlot = function () {
            base.doneCount++;

            if (base.doneCount === base.options.number) {

                if ($.isFunction(base.options.onEnd)) {
                    base.options.onEnd(base.exclude[base.exclude.length - 1]);
                }

                base.isSpinning = false;
            }
        }

        //base.checkWinner = function (endNum, slot) {

        //    base.doneCount++;
        //    // set the slot number to whatever it ended on
        //    slot.number = endNum;

        //    // if its in the winners array
        //    if (
        //        ($.isArray(base.options.winnerNumber) && base.options.winnerNumber.indexOf(endNum) > -1) ||
        //        endNum === base.options.winnerNumber
        //        ) {

        //        // its a winner!
        //        base.winCount++;
        //        base.winners.push(slot.$el);

        //    }

        //    if (base.doneCount === base.options.number) {

        //        var finalNumbers = [];

        //        $.each(base.allSlots, function (index, val) {
        //            finalNumbers[index] = val.number;
        //        });

        //        if ($.isFunction(base.options.onEnd)) {
        //            base.options.onEnd(finalNumbers);
        //        }

        //        if (base.winCount && $.isFunction(base.options.onWin)) {
        //            base.options.onWin(base.winCount, base.winners, finalNumbers);
        //        }
        //        base.isSpinning = false;
        //    }
        //};

        //[2000, 4000, 7000]
        base.playSlots = function (showtime) {

            base.isSpinning = true;
            //base.winCount = 0;
            base.doneCount = 0;
            //base.winners = [];
            var target = base.randomRange(base.options.min, base.options.max, base.exclude);
            base.exclude.push(target);

            // 高位數放在低索引位置上
            for (var i = 1 ; i <= base.options.number ; i++) {
                base.endNum[base.options.number - i] = Math.floor((target % Math.pow(10, i)) * 10 / Math.pow(10, i));
            }

            console.log("target num : " + target);

            // 0 秒不跑動畫
            if (showtime > 0) {
                // 計算 loops 總和
                var totalloops = 0;
                for (var i = 0 ; i < base.options.loops.length ; i++) {
                    totalloops += base.options.loops[i];
                }

                // 計算每個 slot 的 遞增時間
                for (var i = 1 ; i <= base.options.number ; i++) {
                    var time = showtime * base.options.loops[i - 1] / totalloops;

                    base.allSlots[i - 1].increment = (time / base.options.loops[i - 1]) / base.options.loops[i - 1];

                    //console.log("time : (" + i + ") : " + time);
                }                
            }

            // 執行開始前的事件
            if ($.isFunction(base.options.onStart)) {
                base.options.onStart();
            }

            $.each(base.allSlots, function (index, val) {

                if (showtime > 0) {
                    this.spinSpeed = 0;
                    this.loopCount = 0;
                    this.spinEm();
                }
                else {
                    this.noAnimation();
                }
            });

        };

        base.InitialRadomNum = function () {
            base.exclude = [];
        }


        //base.onWin = function () {
        //    if ($.isFunction(base.options.onWin)) {
        //        base.options.onWin();
        //    }
        //};


        // Run initializer
        base.init();
    };


    // --------------------------------------------------------------------- //
    // JQUERY FN
    // --------------------------------------------------------------------- //

    $.fn.jSlots = function (options) {
        if (this.length) {
            return this.each(function () {
                (this.slotmachine = new $.jSlots(this, options));
            });
        }
    };

})(jQuery);
