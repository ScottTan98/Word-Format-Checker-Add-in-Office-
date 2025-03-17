import { OfficeMockObject } from "office-addin-mock";

describe("Word Add-in Tests", () => {
  beforeEach(() => {
    jest.resetModules();

    const OfficeMockData = {
      onReady: jest.fn((callback) => callback({ host: "Word" })),
      HostType: {
        Word: "Word",
      },
    };
    global.Office = new OfficeMockObject(OfficeMockData);

    const WordMockData = {
      context: {
        document: {
          body: {
            paragraphs: {
              items: [],
            },
            getParagraphs: function () {
              return this.paragraphs;
            },
          },
        },
      },
      run: async function (callback) {
        await callback(this.context);
      },
    };
    global.Word = new OfficeMockObject(WordMockData);

    console.log = jest.fn();
  });

  test("No paragraphs will show 'No paragraphs found'", async () => {
    global.Word.context.document.body.paragraphs.items = [];

    const { run } = require("../taskpane/taskpane");
    await run();

    expect(console.log).toHaveBeenCalledWith("No paragraphs found.");
  });

  test("Not enough words will show 'Not enough words in the first paragraph'", async () => {
    global.Word.context.document.body.paragraphs.items = [
      {
        getTextRanges: jest.fn(() => ({
          items: ["word1", "word2"],
          load: jest.fn(),
        })),
        load: jest.fn(),
      },
    ];

    const { run } = require("../taskpane/taskpane");
    await run();

    expect(console.log).toHaveBeenCalledWith("Not enough words in the first paragraph.");
  });

  test("Test bold status of the first word", async () => {
    global.Word.context.document.body.paragraphs.items = [
      {
        getTextRanges: jest.fn(() => ({
          items: [
            { font: { bold: true }, load: jest.fn() },
            { font: { underline: "None" }, load: jest.fn() },
            { font: { size: 12 }, load: jest.fn() },
          ],
          load: jest.fn(),
        })),
        load: jest.fn(),
      },
    ];

    const { run } = require("../taskpane/taskpane");
    await run();

    expect(console.log).toHaveBeenCalledWith("First word bold:", true);
  });

  test("Test underline status of the second word", async () => {
    global.Word.context.document.body.paragraphs.items = [
      {
        getTextRanges: jest.fn(() => ({
          items: [
            { font: { bold: false }, load: jest.fn() },
            { font: { underline: "Single" }, load: jest.fn() },
            { font: { size: 12 }, load: jest.fn() },
          ],
          load: jest.fn(),
        })),
        load: jest.fn(),
      },
    ];

    const { run } = require("../taskpane/taskpane");
    await run();

    expect(console.log).toHaveBeenCalledWith("Second word underlined:", true);
  });

  test("Test font size of the third word", async () => {
    global.Word.context.document.body.paragraphs.items = [
      {
        getTextRanges: jest.fn(() => ({
          items: [
            { font: { bold: false }, load: jest.fn() },
            { font: { underline: "None" }, load: jest.fn() },
            { font: { size: 14 }, load: jest.fn() },
          ],
          load: jest.fn(),
        })),
        load: jest.fn(),
      },
    ];

    const { run } = require("../taskpane/taskpane");
    await run();

    expect(console.log).toHaveBeenCalledWith("Third word font size:", 14);
  });
});