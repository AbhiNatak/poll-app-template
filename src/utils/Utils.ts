import * as uuid from "uuid";

export namespace Utils {
    export let YEARS: string = "YEARS";
    export let MONTHS: string = "MONTHS";
    export let WEEKS: string = "WEEKS";
    export let DAYS: string = "DAYS";
    export let HOURS: string = "HOURS";
    export let MINUTES: string = "MINUTES";
    export let DEFAULT_LOCALE: string = "en";

    export function isValidJson(json: string): boolean {
        try {
            JSON.parse(JSON.stringify(json));
            return true;
        } catch (e) {
            return false;
        }
    }

    export function isEmptyString(str: string): boolean {
        return isEmptyObject(str);
    }

    export function isEmptyObject(obj: any): boolean {
        if (obj == undefined || obj == null) {
            return true;
        }

        var isEmpty = false;

        if (typeof obj === "number" || typeof obj === "boolean") {
            isEmpty = false;
        } else if (typeof obj === "string") {
            isEmpty = obj.trim().length == 0;
        } else if (Array.isArray(obj)) {
            isEmpty = obj.length == 0;
        } else if (typeof obj === "object") {
            if (isValidJson(obj)) {
                isEmpty = JSON.stringify(obj) == "{}";
            }
        }
        return isEmpty;
    }

    export function getTimeRemaining(deadLineDate: Date): {} {
        var now = new Date().getTime();
        var deadLineTime = deadLineDate.getTime();

        var minutes: number = 0;
        var hours: number = 0;
        var days: number = 0;
        var weeks: number = 0;
        var months: number = 0;
        var years: number = 0;

        var diff = Math.abs(deadLineTime - now);
        if (diff > 0) {
            var minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
            var hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));

            var days = Math.floor(
                (diff % (1000 * 60 * 60 * 24 * 7)) / (1000 * 60 * 60 * 24)
            );
            var weeks = Math.floor(
                (diff % (1000 * 60 * 60 * 24 * 30)) / (1000 * 60 * 60 * 24 * 7)
            );
            var months = Math.floor(
                (diff % (1000 * 60 * 60 * 24 * 365)) / (1000 * 60 * 60 * 24 * 30)
            );
            var years = Math.floor(diff / (1000 * 60 * 60 * 24 * 365));
        }
        return {
            YEARS: years,
            MONTHS: months,
            WEEKS: weeks,
            DAYS: days,
            HOURS: hours,
            MINUTES: minutes,
        };
    }

    export function getDefaultExpiry(activeDays: number): Date {
        let date: Date = new Date();
        date.setDate(date.getDate() + activeDays);

        // round off to next 30 minutes time multiple
        if (date.getMinutes() > 30) {
            date.setMinutes(0);
            date.setHours(date.getHours() + 1);
        } else {
            date.setMinutes(30);
        }
        return date;
    }

    export function generateGUID(): string {
        return uuid.v4();
    }

    export function downloadContent(fileName: string, data: string) {
        if (data && fileName) {
            var a = document.createElement("a");
            a.href = data;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }
    }

    export function isRTL(locale: string): boolean {
        let rtlLang: string[] = ["ar", "he", "fl"];
        if (locale && rtlLang.indexOf(locale.split("-")[0]) !== -1) {
            return true;
        } else {
            return false;
        }
    }

    export function announceText(text: string) {
        let ariaLiveSpan: HTMLSpanElement = document.getElementById(
            "aria-live-span"
        );
        if (ariaLiveSpan) {
            ariaLiveSpan.innerText = text;
        } else {
            ariaLiveSpan = document.createElement("SPAN");
            ariaLiveSpan.style.cssText =
                "position: fixed; overflow: hidden; width: 0px; height: 0px;";
            ariaLiveSpan.id = "aria-live-span";
            ariaLiveSpan.innerText = "";
            ariaLiveSpan.setAttribute("aria-live", "polite");
            ariaLiveSpan.tabIndex = -1;
            document.body.appendChild(ariaLiveSpan);
            setTimeout(() => {
                ariaLiveSpan.innerText = text;
            }, 50);
        }
    }

    export function getNonNullString(str: string): string {
        if (isEmptyObject(str)) {
            return "";
        } else {
            return str;
        }
    }
}
