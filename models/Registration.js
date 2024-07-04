// models/Registration.js
import mongoose from 'mongoose';

const registrationSchema = new mongoose.Schema({
  name: String,
  email: String,
  phone: String,
  passOutYear: Number,
  uniqueId: String,
  transactionID: String,
  entered: { type: Boolean, default: false }
});

const Registration = mongoose.model('Registration', registrationSchema);
export default Registration;
