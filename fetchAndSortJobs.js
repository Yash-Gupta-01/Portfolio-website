const fs = require("node:fs/promises");
const path = require("node:path");

const API_URL = "https://6figr.com/api/v1/jobs/search";
const OUTPUT_FILE = path.join(__dirname, "jobs.json");

const requestBody = {
  filters: {
    titles: ["software"],
    skills: [],
    excludedKeywords: [],
    companies: [],
    excludedCompanies: [],
    locations: [],
    remote: false,
    expLevels: ["EN"],
    industries: [],
    fundingStages: [],
    companySize: [],
    jobFamilies: [],
    companyTags: [],
    dateOffset: null,
    baseSalaryMin: null,
    workAuthorization: [],
    roleType: null,
    workMode: [],
    workArrangement: ["Full Time"],
    minimumDegree: null,
    relevancy: null,
    countries: ["IN"],
  },
  page: 1,
  limit: 50,
  includeFilters: true,
  description: false,
  src: "jobs",
};

async function fetchAndSortJobs() {
  const headers = {
    accept: "application/json, text/plain, */*",
    "accept-language": "en-GB,en-US;q=0.9,en;q=0.8",
    "content-type": "application/json",
    "sec-ch-ua": '"Chromium";v="148", "Google Chrome";v="148", "Not/A)Brand";v="99"',
    "sec-ch-ua-mobile": "?1",
    "sec-ch-ua-platform": '"Android"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    referer: "https://6figr.com/jobs-search?titles=software&expLevels=EN",
  };

  // Optional: provide cookie via environment variable if the endpoint expects session cookies.
  if (process.env.SIXFIGR_COOKIE) {
    headers.cookie = process.env.SIXFIGR_COOKIE;
  }

  const response = await fetch(API_URL, {
    method: "POST",
    headers,
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Request failed (${response.status}): ${errorText}`);
  }

  const json = await response.json();
  const jobs = json?.data?.jobs;

  if (!Array.isArray(jobs)) {
    throw new Error("Unexpected response format: data.jobs is not an array");
  }

  const result = jobs
    .map((job) => ({
      postedDate: job?.postedDate ?? null,
      url: job?.url ?? null,
      title: job?.title ?? null,
      name: job?.company?.name ?? null,
    }))
    .sort((a, b) => {
      const timeA = a.postedDate ? Date.parse(a.postedDate) : 0;
      const timeB = b.postedDate ? Date.parse(b.postedDate) : 0;
      return timeB - timeA;
    });

  return result;
}

if (require.main === module) {
  fetchAndSortJobs()
    .then(async (result) => {
      await fs.writeFile(OUTPUT_FILE, JSON.stringify(result, null, 2), "utf8");
      console.log(`Saved ${result.length} jobs to ${OUTPUT_FILE}`);
      console.log(JSON.stringify(result, null, 2));
    })
    .catch((error) => {
      console.error(error.message);
      process.exitCode = 1;
    });
}

module.exports = { fetchAndSortJobs };
