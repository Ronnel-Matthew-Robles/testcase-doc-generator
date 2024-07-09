import mongoose from 'mongoose';

const JiraTicketHelperSchema = new mongoose.Schema({
    threadId: { type: String, unique: true },
    runId: { type: String, unique: true },
    messageId: { type: String, unique: true },
    userStoryNumber: { type: String, unique: true },
    data: Object,
});

const JiraTicketHelper = mongoose.model('JiraTicketHelper', JiraTicketHelperSchema);

export default JiraTicketHelper;