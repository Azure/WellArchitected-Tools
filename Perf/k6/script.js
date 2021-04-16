import http from "k6/http";
import { sleep, check, group } from "k6";
import { Rate, Trend } from "k6/metrics";

//
// Setup custom metrics
// Trends calculate different statistics over multiple requests on the added values (min, max, average, or percentiles)
// Rates represent customer metrics for keeping track of the percentage of added values that are non-zero
//
const testK6PageDurationTrend = new Trend("Test k6 Page Duration");
const testK6PageErrorRate = new Rate("Test k6 Page Errors");
const testK6ApisDurationTrend = new Trend("Test k6 APIs Duration");
const testK6ApisErrorRate = new Rate("Test k6 APIs Errors");
const testRandomPeopleApiDurationTrend = new Trend("Test Random People API Duration");
const testRandomPeopleApiErrorRate = new Rate("Test Random People API Errors");

//
// Options for the performance test run
//
export const options = {
  //
  // Stages allow you to specify ramp up/down patters for your tests.
  // In the below pattern, requests ramp up to 30 VUs (virtual users) in 3 stages over 3 minutes.
  // Then, the requests ramp back down to 0 VUs in 2 stages over 2 minutes.
  //
  stages: [
    { duration: "1m", target: 10 },
    { duration: "1m", target: 20 },
    { duration: "1m", target: 30 },
    { duration: "1m", target: 15 },
    { duration: "1m", target: 0 },
  ],

  //
  // Define "allowable" thresholds
  // In the below example...
  //    - the total duration of the http request should not exceed 10 seconds (e.g. max<10000)
  //    - under 90% of the targeted load (30 VUs * .9 = 27 VUs), the http request duration should be under 2 seconds (e.g. p(90)<2000)
  //    - under 95% of the targeted load (30 VUs * .95 = 29 VUs), the http request duration should be under 5 seconds (e.g. p(95)<5000), understanding that more load may have some adverse impact on performance
  //    - of all requests, less than 10% should fail
  //
  thresholds: {
    "http_req_duration": ["max<10000", "p(90)<2000", "p(95)<5000"],
    "failed requests": ["rate<0.1"],
  },
};

//
// Example of testing the loading of a static page or asset (e.g. css, javascript, image, etc.)
//
function testK6Page() {
  const start = new Date();

  const response = http.get("http://test.k6.io");

  const end = new Date();

  // Capture the time it took to load the entire page
  testK6PageDurationTrend.add(end - start);

  // Make sure that the request returned successfully (e.g. HTTP status of 200). If not, add to the error count.
  check(response, {
    "testK6Page: status is 200": (r) => r.status === 200,
  }) || testK6PageErrorRate.add(1);
}

//
// Example of testing the loading of multiple APIs (e.g. a data-driven page may require multiple APIs for the page to fully render)
//
function testK6Apis() {
  // Define a "group" of endpoints to represent a data-driven page
  group('Sample Dashboard', () => {
    const start = new Date();

    const responses = http.batch([
        ["GET", "https://test-api.k6.io/public/crocodiles/1/"],
        ["GET", "https://test-api.k6.io/public/crocodiles/2/"],
        ["GET", "https://test-api.k6.io/public/crocodiles/3/"],
        ["GET", "https://test-api.k6.io/public/crocodiles/4/"],
        ["GET", "https://test-api.k6.io/public/crocodiles/5/"],
    ]);

    const end = new Date();

    // Capture the time it took to load all api data
    testK6ApisDurationTrend.add(end - start);

    // Make sure that all api requests returned successfully (e.g. HTTP status of 200). If not, add to the error count.
    responses.forEach((response) => {
        check(response, {
        "testK6Apis: status is 200": (r) => r.status === 200,
        }) || testK6ApisErrorRate.add(1);
    });
  });
}

//
// Example of testing the **overloading** (e.g. reaching threshold of allowed requests) of a given API
//
function testRandomPeopleApi() {
  const start = new Date();

  const response = http.get("https://randomuser.me/api/?results=200");

  const end = new Date();

  // Capture the time it took to load all api data
  testRandomPeopleApiDurationTrend.add(end - start);

  // Make sure that all api requests returned successfully (e.g. HTTP status of 200). If not, add to the error count.
  check(response, {
    "testRandomPeopleApi: status is 200": (r) => r.status === 200,
  }) || testRandomPeopleApiErrorRate.add(1);
}

//
// This is the main entry point for tests
//
export default function () {
  testK6Page();
  testK6Apis();
  //testRandomPeopleApi();

  // wait 1 second before another request in order to simulate clicks of a real user (and not create false positives for DDoS)
  sleep(1);
}
