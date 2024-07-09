import mongoose from 'mongoose';

const FileReferenceSchema = new mongoose.Schema({
    userStoryNumber: { type: String, required: true },
    openAIFileId: { type: String, required: true },
    filename: { type: String, required: true },
    bytes: { type: Number, required: true },
});

const FileReference = mongoose.model('FileReference', FileReferenceSchema);

export default FileReference;