import mongoose from 'mongoose';

const QATestGeneratorSchema = new mongoose.Schema({
    threadId: { type: String, unique: true },
    runId: { type: String, unique: true },
    messageId: { type: String, unique: true },
    userStoryNumber: { type: String, unique: true },
    title: { type: String },
    data: Object,
    testIssues: Array,
});

const QATestGenerator = mongoose.model('QATestGenerator', QATestGeneratorSchema);

export default QATestGenerator;